Imports System.Data
Imports System.Data.Common
Imports CableSoft.BLL.Utility
Imports CableSoft.Utility.DataAccess
Imports CounterPayLanguage = CableSoft.BLL.Language.SO61.CounterPayLanguage

Public Class CounterPay
    Inherits CableSoft.BLL.Utility.BLLBasic
    Implements IDisposable
    'Implements CableSoft.SL.Printing.Web.IReportData

    Private _DAL As New CounterPayDALMultiDB(Me.LoginInfo.Provider)

#Region "New"
    Public Sub New()
        MyBase.New()
    End Sub
    Public Sub New(ByVal LoginInfo As CableSoft.BLL.Utility.LoginInfo)
        MyBase.New(LoginInfo)
    End Sub
    Public Sub New(ByVal LoginInfo As LoginInfo, ByVal DBConnection As System.Data.Common.DbConnection)
        MyBase.New(LoginInfo, DBConnection)
    End Sub
    Public Sub New(ByVal LoginInfo As LoginInfo, ByVal DAO As CableSoft.Utility.DataAccess.DAO)
        MyBase.New(LoginInfo, DAO)
    End Sub
#End Region

    '取得畫面元件資料
    Public Function OpenAllData() As DataSet
        Dim ds As New DataSet
        Dim GroupId As String = "0"
        If String.IsNullOrEmpty(Me.LoginInfo.GroupId) = False Then GroupId = Me.LoginInfo.GroupId
        'CD039
        Using dt As DataTable = DAO.ExecQry(_DAL.GetCompCode(GroupId, Me.LoginInfo.EntryId))
            dt.TableName = "CompCode"
            ds.Tables.Add(dt.Copy)
        End Using
        'CD031
        Using dt As DataTable = DAO.ExecQry(_DAL.GetCMCode)
            dt.TableName = "CMCode"
            ds.Tables.Add(dt.Copy)
        End Using
        'CD032
        Using dt As DataTable = DAO.ExecQry(_DAL.GetPTCode)
            dt.TableName = "PTCode"
            ds.Tables.Add(dt.Copy)
        End Using
        'CM003
        Using dt As DataTable = DAO.ExecQry(_DAL.GetClctEn)
            dt.TableName = "ClctEn"
            ds.Tables.Add(dt.Copy)
        End Using
        Using Utility As New CableSoft.SO.BLL.Utility.Utility(Me.LoginInfo, DAO)
            'SO029
            Using dt As DataTable = Utility.GetPriv(LoginInfo.EntryId, New String() {"SO3318", "SO33181", "SO33182", "SO33183", "SO33184", "SO33185", "SO331821", "SO331851", "SO114FA5"})
                dt.TableName = "Privs"
                ds.Tables.Add(dt.Copy)
            End Using
            'SO041
            Using dt As DataTable = Utility.GetSystem(BLL.Utility.SystemTableType.System, "*", "")
                dt.TableName = "System"
                ds.Tables.Add(dt.Copy)
            End Using
            'SO043
            Using dt As DataTable = Utility.GetSystem(BLL.Utility.SystemTableType.Charge, "*", "")
                dt.TableName = "SystemCharge"
                ds.Tables.Add(dt.Copy)
            End Using
        End Using
        'CD125
        Using dt As DataTable = DAO.ExecQry(_DAL.GetInvConSetting, New String() {LoginInfo.CompCode})
            dt.TableName = "InvConSetting"
            ds.Tables.Add(dt.Copy)
        End Using
        'SO062
        Using dt As DataTable = DAO.ExecQry(_DAL.GetTranDate, New String() {LoginInfo.CompCode})
            dt.TableName = "TranDate"
            ds.Tables.Add(dt.Copy)
        End Using
        'CD122
        Using dt As DataTable = DAO.ExecQry(_DAL.GetCarrierType)
            dt.TableName = "CarrierType"
            ds.Tables.Add(dt.Copy)
        End Using
        'CD037
        Using dt As DataTable = DAO.ExecQry(_DAL.GetCardCode)
            dt.TableName = "CardCode"
            ds.Tables.Add(dt.Copy)
        End Using

        Return ds
    End Function
    '取得Grid顯示收費
    Public Function GetChargeTmp(ByVal Para As DataTable) As DataSet
        Dim ds As New DataSet
        Dim strWhere As String = " A.CompCode= " & LoginInfo.CompCode
        Dim intTotalAmt As Integer = 0, strMaxEntryNo As String = "", intCountNo As Integer = 0

        If Para.Rows(0).IsNull("ClctEn") = False Then
            strWhere = String.Format("{0} And A.RealDate=To_Date('{1}','yyyy/MM/dd') And A.ClctEn='{2}' ", strWhere, Para.Rows(0).Item("RealDate").ToString, Para.Rows(0).Item("ClctEn").ToString)
        Else
            strWhere = String.Format("{0} And A.RealDate=To_Date('{1}','yyyy/MM/dd')", strWhere, Para.Rows(0).Item("RealDate").ToString)
        End If

        Using dtShow As DataTable = DAO.ExecQry(_DAL.GetShowCount(strWhere))
            If dtShow.Rows.Count > 0 Then
                intTotalAmt = Integer.Parse(dtShow.Rows(0).Item("TotalAmt"))
                If dtShow.Rows(0).IsNull("MaxEntryNo") = False Then strMaxEntryNo = dtShow.Rows(0).Item("MaxEntryNo").ToString
                intCountNo = Integer.Parse(dtShow.Rows(0).Item("CountNo"))
            End If
            Using dtTmp As DataTable = DAO.ExecQry(_DAL.GetChargeTmp(strWhere, intTotalAmt, strMaxEntryNo, intCountNo))
                dtTmp.TableName = "ChargeTmp"
                ds.Tables.Add(dtTmp.Copy)
            End Using
        End Using

        Return ds
    End Function
    '回傳資料到前端檢核
    Public Function ChkDataOK(ByVal strBillNo As String) As CableSoft.BLL.Utility.RIAResult
        Try
            Dim strBillSQL As String = ""
            Dim intErrCode As Integer = 1
            Dim blnData As Boolean = False

            Select Case strBillNo.Length
                Case 11
                    strBillSQL = String.Format("A.MediaBillNo='{0}'", strBillNo)
                Case 12
                    strBillSQL = String.Format("A.PrtSNo='{0}'", strBillNo)
                Case 15
                    strBillSQL = String.Format("A.BillNo='{0}'", strBillNo)
                Case 16
                    strBillSQL = String.Format("A.MediaBillNo='{0}'", strBillNo)
                Case Else
                    '2019.03.15 舊版不會出現檢核提示訊息 找不到就直接沒資料
                    'Return New RIAResult With {.ResultBoolean = False, .ErrorCode = -100, .ErrorMessage = CounterPayLanguage.ChkBillOK}
            End Select

            Dim ds As New DataSet
            '取得收費資料SO033
            Dim dtCharge As DataTable = DAO.ExecQry(_DAL.GetRealCharge2(Me.LoginInfo.CompCode, strBillSQL))
            dtCharge.TableName = "RealCharge"
            ds.Tables.Add(dtCharge.Copy)
            '取得週期性收費資料SO003
            Dim dtPeriod As DataTable = DAO.ExecQry(_DAL.GetPeriodData(Me.LoginInfo.CompCode, strBillSQL))
            dtPeriod.TableName = "PeriodCycle"
            ds.Tables.Add(dtPeriod.Copy)
            'SO043.Para41
            Dim dtChargeSystem As DataTable = DAO.ExecQry(_DAL.GetSO043())
            dtChargeSystem.TableName = "ChargeSystem"
            ds.Tables.Add(dtChargeSystem.Copy)

            '2019.03.15 如果入T單他KEY到11,12碼的時候會跳出一個無單據的訊息(照舊版拿掉檢核)
            If dtCharge IsNot Nothing Then
                If dtCharge.Rows.Count > 0 Then
                    Return New RIAResult With {.ResultBoolean = True, .ErrorCode = 0, .ErrorMessage = "", .ResultDataSet = ds.Copy}
                End If
            End If
            Return New RIAResult With {.ResultBoolean = blnData, .ErrorCode = intErrCode, .ErrorMessage = ""}
        Catch ex As Exception
            Return New RIAResult With {.ResultBoolean = False, .ErrorCode = -1, .ErrorMessage = ex.ToString()}
        End Try
    End Function

    '新增櫃檯繳款費用
    Public Function AddNewCharge(ByVal strBillNo As String, ByVal dtPara As DataTable) As CableSoft.BLL.Utility.RIAResult
        Dim trans As DbTransaction = Nothing
        Dim MyTrans As Boolean
        Dim cn As System.Data.Common.DbConnection = DAO.GetConn()
        Try
            Dim strBillSQL As String = "", strErrMsg As String = ""
            Select Case strBillNo.Length
                Case 11
                    strBillSQL = String.Format("A.MediaBillNo='{0}'", strBillNo)
                Case 12
                    strBillSQL = String.Format("A.PrtSNo='{0}'", strBillNo)
                Case 15
                    strBillSQL = String.Format("A.BillNo='{0}'", strBillNo)
                Case 16
                    strBillSQL = String.Format("A.MediaBillNo='{0}'", strBillNo)
                Case Else
                    Return New RIAResult With {.ResultBoolean = False, .ErrorCode = -100, .ErrorMessage = CounterPayLanguage.ChkBillOK}
            End Select
            '取得收費資料
            Dim dtCharge As DataTable = DAO.ExecQry(_DAL.GetRealCharge(Me.LoginInfo.CompCode, strBillSQL))
            If dtCharge.Rows.Count = 0 Then
                Return New RIAResult With {.ResultBoolean = False, .ErrorCode = -199, .ErrorMessage = CounterPayLanguage.ChkBillOK}
            End If
            dtCharge.TableName = "RealCharge"
            '檢核收費是否可做入帳
            If IsDataOk(strBillSQL, dtCharge.Rows(0).Item("Custid"), dtCharge.Rows(0).Item("ServiceType").ToString, strErrMsg) = False Then
                Return New RIAResult With {.ResultBoolean = False, .ErrorCode = -199, .ErrorMessage = strErrMsg}
            End If
            If DAO.Transaction Is Nothing Then
                MyTrans = True
                If cn.State = ConnectionState.Closed Then
                    cn.ConnectionString = Me.LoginInfo.ConnectionString
                    cn.Open()
                End If

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
            'CableSoft.BLL.Utility.Utility.SetClientInfo(DAO, LoginInfo.EntryId)
            If MyTrans Then CableSoft.BLL.Utility.Utility.SetClientInfo(DAO, LoginInfo.EntryId, CounterPayLanguage.ProcessKind1)

            '新增收費到SO074A
            If EntNewCharge(dtPara, dtCharge, strErrMsg) = False Then
                Return New RIAResult With {.ResultBoolean = False, .ErrorCode = -189, .ErrorMessage = strErrMsg}
            End If

            '重取SO074A顯示到畫面
            Dim ds As DataSet = GetChargeTmp(dtPara)

            If MyTrans Then
                trans.Commit()
                'cn.Close()
                'CableSoft.BLL.Utility.Utility.ClearClientInfo(DAO)
                'DAO.AutoCloseConn = True
                'DAO.Dispose()
                'cn.Dispose()
            End If

            Return New RIAResult With {.ResultBoolean = True, .ErrorCode = 0, .ErrorMessage = strErrMsg, .ResultDataSet = ds.Copy}
        Catch ex As Exception
            If MyTrans Then trans.Rollback()
            Return New RIAResult With {.ErrorCode = -1, .ErrorMessage = "AddNewCharge Error!!" & ex.ToString(), .ResultBoolean = False}
        Finally
            If MyTrans Then
                CableSoft.BLL.Utility.Utility.ClearClientInfo(DAO)
                'trans.Commit()
                cn.Close()
                DAO.AutoCloseConn = True
                DAO.Dispose()
                cn.Dispose()
            End If
        End Try
    End Function
    '檢核是否可臨櫃入帳
    Private Function IsDataOk(ByVal strBillSQL As String, ByVal Custid As Integer, ByVal ServiceType As String, ByRef strErrMsg As String) As Boolean
        IsDataOk = False
        'SO074A
        Using dt As DataTable = DAO.ExecQry(_DAL.GetTmpCharge(LoginInfo.CompCode, strBillSQL))
            If dt.Rows.Count > 0 Then
                strErrMsg = CounterPayLanguage.BillOver
                Return False
            End If
        End Using
        '客戶狀態  
        Using dt As DataTable = DAO.ExecQry(_DAL.GetCustStatus(Custid, ServiceType))
            If dt.Rows.Count > 0 Then
                Dim intCustStatusCode As Integer = dt.Rows(0).Item("CustStatusCode")
                If intCustStatusCode <> 1 Then
                    strErrMsg = CounterPayLanguage.CustNotOk
                    'Return False
                End If
            End If
        End Using

        IsDataOk = True

        Return IsDataOk
    End Function
    '新增臨櫃入帳收費
    Public Function EntNewCharge(ByVal dtPara As DataTable, ByVal dtCharge As DataTable, ByRef strErrMsg As String, Optional ByVal blnFromEPG As Boolean = False) As Boolean
        EntNewCharge = False
        Dim intMaxEntryNo As Integer = 0
        '是否有傳入參數&收費資料
        If dtPara Is Nothing Then
            strErrMsg = CounterPayLanguage.NoParaData
            Return False
        End If
        If dtCharge Is Nothing Then
            strErrMsg = CounterPayLanguage.NoChargeData
            Return False
        End If
        '取得序號
        Using dtMax As DataTable = DAO.ExecQry(_DAL.GetMaxEntryNo(dtCharge.Rows(0).Item("CompCode"), dtPara.Rows(0).Item("ClctEn").ToString, dtPara.Rows(0).Item("RealDate").ToString))
            intMaxEntryNo = Integer.Parse(dtMax.Rows(0).Item(0)) + 1
        End Using
        '新增SO074A
        If AddChargeTmp(dtPara, dtCharge, intMaxEntryNo, blnFromEPG) = False Then
            strErrMsg = CounterPayLanguage.AddChargeTmpErr
            Return False
        End If
        '更新SO033
        If UpdRealCharge(dtPara, dtCharge, blnFromEPG) = False Then
            strErrMsg = CounterPayLanguage.UpdRealChargeErr
            Return False
        End If

        EntNewCharge = True

        Return EntNewCharge
    End Function
    '新增SO074A
    Private Function AddChargeTmp(ByVal dtPara As DataTable, ByVal dtCharge As DataTable, ByVal intMaxEntryNo As Integer, ByVal blnFromEPG As Boolean) As Boolean
        Dim bll As New CableSoft.SO.BLL.Billing.Utility.Utility(LoginInfo, DAO)
        Dim obj As New CableSoft.SO.BLL.Utility.Utility(LoginInfo, DAO)
        Dim CSLog As CableSoft.SO.BLL.DataLog.DataLog = Nothing
        CSLog = New CableSoft.SO.BLL.DataLog.DataLog(LoginInfo, DAO)

        'CustName
        Dim dtCustData As DataTable = DAO.ExecQry(_DAL.GetCustName(LoginInfo.CompCode, dtCharge.Rows(0).Item("Custid")))
        'SO074A
        Dim dtChargeTmp As DataTable = DAO.ExecQry(_DAL.GetTmpCharge(LoginInfo.CompCode, "A.BillNo='' And A.Item=-1"))
        For Each ChargeRow As DataRow In dtCharge.Rows
            Dim drRow As DataRow = dtChargeTmp.NewRow
            dtChargeTmp.Rows.Add(drRow)
            CableSoft.BLL.Utility.Utility.CopyDataRow(ChargeRow, drRow)

            drRow.Item("CustName") = dtCustData.Rows(0).Item(0).ToString
            drRow.Item("RcdRowID") = ChargeRow.Item("CTID").ToString
            'If blnFromEPG Then
            drRow.Item("SUCCode") = ChargeRow.Item("UCCode")
            drRow.Item("SUCName") = ChargeRow.Item("UCName").ToString
            'Else
            '    drRow.Item("SUCCode") = dtPara.Rows(0).Item("UCCode")
            '    drRow.Item("SUCName") = dtPara.Rows(0).Item("UCName").ToString
            'End If
            drRow.Item("EntryNO") = intMaxEntryNo
            drRow.Item("RealDate") = dtPara.Rows(0).Item("RealDate")
            drRow.Item("EntryEn") = Me.LoginInfo.EntryId
            drRow.Item("CMCode") = dtPara.Rows(0).Item("CMCode")
            drRow.Item("CMName") = dtPara.Rows(0).Item("CMName").ToString
            drRow.Item("ClctEn") = dtPara.Rows(0).Item("ClctEn").ToString
            drRow.Item("ClctName") = dtPara.Rows(0).Item("ClctName").ToString
            drRow.Item("PTCode") = dtPara.Rows(0).Item("PTCode").ToString
            drRow.Item("PTName") = dtPara.Rows(0).Item("PTName").ToString
            drRow.Item("RealAmt") = ChargeRow.Item("ShouldAmt")
            drRow.Item("UPDEN") = Me.LoginInfo.EntryName
            drRow.Item("UPDTIME") = DateTimeUtility.GetDTString(DateTime.Now)
        Next

        Dim aDtClone As DataTable = dtChargeTmp.Copy

        For intLoop As Integer = 0 To aDtClone.Rows.Count - 1
            Dim CloseRow As DataRow = aDtClone.Rows(intLoop)
            If Not CableSoft.BLL.Utility.Utility.ExecuteCommand(DAO, UpdateMode.InsertRow, aDtClone, intLoop, "SO074A",
                String.Format("BillNo = '{0}' And Item = {1}", CloseRow.Item("BillNo"), CloseRow.Item("Item")), , , False) Then
                Return False
            End If
        Next

        Dim aResult As RIAResult = CSLog.SummaryExpansion(DataLog.OpType.Update, "SO074A", aDtClone, Int32.Parse(Integer.Parse(DateTime.Now.ToString("yyyyMMdd"))))
        If Not aResult.ResultBoolean Then
            Select Case aResult.ErrorCode
                Case -6
                    Return False
                Case Else
            End Select
        End If
        If bll IsNot Nothing Then bll.Dispose()
        If obj IsNot Nothing Then obj.Dispose()

        Return True
    End Function
    '更新SO033
    Private Function UpdRealCharge(ByVal dtPara As DataTable, ByVal dtCharge As DataTable, ByVal blnFromEPG As Boolean) As Boolean
        UpdRealCharge = False

        Dim CSLog As CableSoft.SO.BLL.DataLog.DataLog = Nothing
        CSLog = New CableSoft.SO.BLL.DataLog.DataLog(LoginInfo, DAO)

        'CD013
        Dim dtUCCode As DataTable = DAO.ExecQry(_DAL.GetUCCode(IIf(blnFromEPG = True, " RefNo=8", " RefNo=3")))
        If dtCharge.Columns.Contains("ctid") Then
            'dtCharge.Columns.Remove(dtCharge.Columns("ctid"))
        End If
        If dtCharge.Columns.Contains("CTID") Then
            'dtCharge.Columns.Remove(dtCharge.Columns("CTID"))
        End If


        For intLoop As Integer = 0 To dtCharge.Rows.Count - 1

            Dim CloseRow As DataRow = dtCharge.Rows(intLoop)
            Dim Affected As Integer = 0

            CloseRow.Item("UCCode") = dtUCCode.Rows(0).Item("CodeNo")
            CloseRow.Item("UCName") = dtUCCode.Rows(0).Item("Description").ToString
            CloseRow.Item("RealDate") = dtPara.Rows(0).Item("RealDate")
            CloseRow.Item("RealAmt") = 0
            CloseRow.Item("CMCode") = dtPara.Rows(0).Item("CMCode")
            CloseRow.Item("CMName") = dtPara.Rows(0).Item("CMName").ToString
            CloseRow.Item("ClctEn") = dtPara.Rows(0).Item("ClctEn").ToString
            CloseRow.Item("ClctName") = dtPara.Rows(0).Item("ClctName").ToString
            CloseRow.Item("PTCode") = dtPara.Rows(0).Item("PTCode").ToString
            CloseRow.Item("PTName") = dtPara.Rows(0).Item("PTName").ToString
            CloseRow.Item("UPDEN") = Me.LoginInfo.EntryName
            CloseRow.Item("UPDTIME") = DateTimeUtility.GetDTString(DateTime.Now)
            CloseRow.Item("NEWUPDTIME") = DateTime.Now

            If Not CableSoft.BLL.Utility.Utility.ExecuteCommand(DAO, UpdateMode.UpdateRow, dtCharge, intLoop, "SO033",
                String.Format("BillNo='{0}' And Item={1}", CloseRow.Item("BillNo"), CloseRow.Item("Item")), , Affected, False) Then
                Return False
            End If
        Next

        Dim aResult As RIAResult = CSLog.SummaryExpansion(DataLog.OpType.Update, "SO033", dtCharge, Int32.Parse(Integer.Parse(DateTime.Now.ToString("yyyyMMdd"))))
        If Not aResult.ResultBoolean Then
            Select Case aResult.ErrorCode
                Case -6
                    Return False
                Case Else
            End Select
        End If

        UpdRealCharge = True

        Return UpdRealCharge
    End Function
    '更新SO074A
    Public Function UpdChargeTmp(ByVal dsCharge As DataSet) As CableSoft.BLL.Utility.RIAResult
        Dim trans As DbTransaction = Nothing
        Dim MyTrans As Boolean
        Dim cn As System.Data.Common.DbConnection = DAO.GetConn()

        Dim CSLog As CableSoft.SO.BLL.DataLog.DataLog = Nothing
        CSLog = New CableSoft.SO.BLL.DataLog.DataLog(LoginInfo, DAO)

        Dim dtCharge As DataTable = dsCharge.Tables("Simple").Copy
        Dim dtPara As DataTable = dsCharge.Tables("Para")
        Dim strWhere As String = ""

        Try
            If DAO.Transaction Is Nothing Then
                MyTrans = True
                If cn.State = ConnectionState.Closed Then
                    cn.ConnectionString = Me.LoginInfo.ConnectionString
                    cn.Open()
                End If

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
            'CableSoft.BLL.Utility.Utility.SetClientInfo(DAO, LoginInfo.EntryId)
            If MyTrans Then CableSoft.BLL.Utility.Utility.SetClientInfo(DAO, LoginInfo.EntryId, CounterPayLanguage.ProcessKind2)

            strWhere = String.Format("A.BillNo='{0}' And A.Item={1}", dtCharge.Rows(0).Item("BillNo").ToString, dtCharge.Rows(0).Item("Item"))

            'SO074A
            Dim dtChargeTmp As DataTable = DAO.ExecQry(_DAL.GetTmpCharge(LoginInfo.CompCode, strWhere))

            Dim ChargeRow As DataRow = dtCharge.Rows(0)  '前端傳入
            Dim TmpRow As DataRow = dtChargeTmp.Rows(0) 'SO074A
            Dim Affected As Integer = 0

            CableSoft.BLL.Utility.Utility.CopyDataRow(ChargeRow, TmpRow)

            If Not CableSoft.BLL.Utility.Utility.ExecuteCommand(DAO, UpdateMode.UpdateRow, dtChargeTmp, 0, "SO074A",
                String.Format("BillNo='{0}' And Item={1}", TmpRow.Item("BillNo"), TmpRow.Item("Item")), , Affected, False) Then
                Return New RIAResult With {.ResultBoolean = False, .ErrorCode = -1, .ErrorMessage = ""}
            End If

            Dim aResult As RIAResult = CSLog.SummaryExpansion(DataLog.OpType.Update, "SO074A", dtCharge, Int32.Parse(Integer.Parse(DateTime.Now.ToString("yyyyMMdd"))))
            If Not aResult.ResultBoolean Then
                Select Case aResult.ErrorCode
                    Case -6
                        Return New RIAResult With {.ResultBoolean = False, .ErrorCode = -1, .ErrorMessage = ""}
                    Case Else
                End Select
            End If

            '重取SO074A顯示到畫面
            Dim ds As DataSet = GetChargeTmp(dtPara)

            If MyTrans Then
                trans.Commit()
                'cn.Close()
                'CableSoft.BLL.Utility.Utility.ClearClientInfo(DAO)
                'DAO.AutoCloseConn = True
                'DAO.Dispose()
                'cn.Dispose()
            End If

            Return New RIAResult With {.ResultBoolean = True, .ErrorCode = 0, .ErrorMessage = "", .ResultDataSet = ds.Copy}
        Catch ex As Exception
            If MyTrans Then trans.Rollback()
            Return New RIAResult With {.ErrorCode = -4, .ErrorMessage = ex.ToString(), .ResultBoolean = False}
        Finally
            If MyTrans Then
                CableSoft.BLL.Utility.Utility.ClearClientInfo(DAO)
                'trans.Commit()
                cn.Close()
                DAO.AutoCloseConn = True
                DAO.Dispose()
                cn.Dispose()
            End If
        End Try
    End Function
    '取得收費資料SO033
    Public Function GetSimple(ByVal BillNo As String, ByVal Item As Integer) As DataTable
        Dim dt As DataTable = DAO.ExecQry(_DAL.GetSimple(Me.LoginInfo.CompCode, BillNo, Item))
        dt.TableName = "Simple"
        Dim strWhere As String = " BillNo='" & BillNo & "' And Item=" & Item & ""
        Dim dt2 As DataTable = DAO.ExecQry(_DAL.GetChargeTmp2(Me.LoginInfo.CompCode, strWhere))

        For intLoop As Integer = 0 To dt.Columns.Count - 1
            If dt2.Columns.IndexOf(dt.Columns(intLoop).ColumnName) >= 0 Then
                dt.Rows(0).Item(intLoop) = dt2.Rows(0).Item(dt.Columns(intLoop).ColumnName)
            End If
        Next
        Return dt
    End Function
    '檢核是否可刪除登錄資料
    Private Function ChkCanDelete(ByVal dtData As DataTable, ByRef strErrMsg As String) As Boolean
        ChkCanDelete = False
        strErrMsg = String.Empty

        For Each row As DataRow In dtData.Rows
            Dim dtCharge As DataTable = DAO.ExecQry(_DAL.GetSimple(LoginInfo.CompCode, row.Item("BillNo").ToString, row.Item("Item")))
            '970801 #4010 若該資料已無未收原因且已有實收日期者，請提示【已入實收或作廢】不可取消。
            If dtCharge.Rows.Count > 0 Then
                If dtCharge.Rows(0).IsNull("RealDate") = False And dtCharge.Rows(0).IsNull("UCCode") Then
                    strErrMsg = CounterPayLanguage.ChargePayOK
                    Return False
                End If
                '已作廢不得取消
                If dtCharge.Rows(0).IsNull("CancelFlag") = False Then
                    If dtCharge.Rows(0).Item("CancelFlag") = 1 Then
                        strErrMsg = CounterPayLanguage.ChargePayOK
                        Return False
                    End If
                End If
            End If
            '990114 #5465 若有啟動VOD,先判斷是否可退費--若其為負項（SIGN='-'）且參考號為21,22，在實收金額KEY入後隨即判斷該VODACCOUNTID  '關於VOD點數檢核,暫緩

            ''有信用卡單號,要退刷  暫時不能處理 等整合流程
            'If row.IsNull("CardBillNo") = False Then
            '    strErrMsg = CounterPayLanguage.ChargePaymentOK
            '    Return False
            'End If
            ''961105 #3591 V529 取消登錄時,增加判斷此筆收費資料是否已被作廢,若已作廢則仍取消登錄,但不做資料異動,並秀訊息告知"該筆收費資料已被作廢,不異動該收費資料!!"
            'If row.IsNull("CancelFlag") = False Then
            '    If row.Item("CancelFlag") = 1 Then
            '        strErrMsg = CounterPayLanguage.ChargeCancelOK
            '        'Return False
            '    End If
            'End If
        Next

        ChkCanDelete = True
        Return ChkCanDelete
    End Function
    '刪除登錄收費資料
    Public Function DeleteChargeTmp(ByVal dtPara As DataTable, ByVal EntryNo As Integer, ByVal RealDate As String, ByVal EntryEn As String, ByVal ClctEn As String) As CableSoft.BLL.Utility.RIAResult
        Dim strErrMsg As String = String.Empty

        '950508 調整刪除資料的PK條件EntryNo,EntryEn,RealDate,ClctEn
        Dim strWhere As String = String.Format(" A.RealDate=To_Date('{0}','yyyy/MM/dd') And A.EntryEn='{1}' And A.EntryNo={2} And A.ClctEn='{3}' ", RealDate, EntryEn, EntryNo, ClctEn)
        Dim dtChargeTmp As DataTable = DAO.ExecQry(_DAL.GetTmpCharge(LoginInfo.CompCode, strWhere))
        '從資料庫SO074A中重新取得資料
        If dtChargeTmp.Rows.Count <= 0 Then
            Return New RIAResult With {.ResultBoolean = False, .ErrorCode = -885, .ErrorMessage = CounterPayLanguage.ChargeTmpNoData}
        End If
        '檢核是否可刪除登錄資料
        If ChkCanDelete(dtChargeTmp, strErrMsg) = False Then
            Return New RIAResult With {.ResultBoolean = False, .ErrorCode = -886, .ErrorMessage = strErrMsg}
        End If

        Dim trans As DbTransaction = Nothing
        Dim MyTrans As Boolean = False
        'Dim cn As System.Data.Common.DbConnection = DAO.GetConn()
        Dim cn As DbConnection = DAO.GetConn()
        Try
            'If DAO.Transaction Is Nothing Then
            '    MyTrans = True
            '    cn.ConnectionString = Me.LoginInfo.ConnectionString
            '    cn.Open()
            '    trans = cn.BeginTransaction
            '    DAO.AutoCloseConn = False
            'Else
            '    MyTrans = False
            '    If cn.State = ConnectionState.Closed Then
            '        cn.Open()
            '    End If
            '    trans = DAO.Transaction
            'End If
            If DAO.Transaction IsNot Nothing Then
                trans = DAO.Transaction
            Else
                MyTrans = True
                If cn.State = ConnectionState.Closed Then
                    cn.ConnectionString = Me.LoginInfo.ConnectionString
                    cn.Open()
                End If

                trans = cn.BeginTransaction
                DAO.Transaction = trans

            End If
            DAO.AutoCloseConn = False
            'DAO.Transaction = trans
            'CableSoft.BLL.Utility.Utility.SetClientInfo(DAO, LoginInfo.EntryId)
            If MyTrans Then CableSoft.BLL.Utility.Utility.SetClientInfo(DAO, LoginInfo.EntryId, CounterPayLanguage.ProcessKind3)

            'LOOP DELETE
            For Each row As DataRow In dtChargeTmp.Rows
                Dim strWhere2 As String = String.Format(" A.RealDate=To_Date('{0}','yyyy/MM/dd') And A.EntryEn='{1}' And A.EntryNo={2} And A.ClctEn='{3}' ", RealDate, EntryEn, EntryNo, ClctEn)
                DAO.ExecNqry(_DAL.DeleteChargeTmp(strWhere2))
                If row.Item("CancelFlag") <> 1 Then
                    '更新SO033
                    If UCSO033(row) = False Then
                        Return New RIAResult With {.ResultBoolean = False, .ErrorCode = -887, .ErrorMessage = CounterPayLanguage.UpdRealChargeErr}
                    End If
                End If
            Next

            '961105 #3591 V529 取消登錄時,增加判斷此筆收費資料是否已被作廢,若已作廢則仍取消登錄,但不做資料異動,並秀訊息告知"該筆收費資料已被作廢,不異動該收費資料!!"

            '發票系統啟動  暫緩

            'Refresh
            '重取SO074A顯示到畫面
            Dim ds As DataSet = GetChargeTmp(dtPara)

            If MyTrans Then
                trans.Commit()
                'cn.Close()
                'CableSoft.BLL.Utility.Utility.ClearClientInfo(DAO)
                'DAO.AutoCloseConn = True
                'DAO.Dispose()
                'cn.Dispose()
            End If

            Return New RIAResult With {.ResultBoolean = True, .ErrorCode = 0, .ErrorMessage = strErrMsg, .ResultDataSet = ds.Copy}
        Catch ex As Exception
            If MyTrans Then trans.Rollback()
            Return New RIAResult With {.ErrorCode = -2, .ErrorMessage = "DeleteChargeTmp Error!!" & ex.ToString(), .ResultBoolean = False}
        Finally
            If MyTrans Then
                CableSoft.BLL.Utility.Utility.ClearClientInfo(DAO)
                'trans.Commit()
                cn.Close()
                DAO.AutoCloseConn = True
                DAO.Dispose()
                cn.Dispose()
            End If
        End Try
    End Function
    '取得原收費資料設定..刪除登錄檔後回填收費資料檔
    Private Function UCSO033(ByVal drData As DataRow) As Boolean
        UCSO033 = False
        Dim strCMCode As String = String.Empty, strCMName As String = String.Empty
        Dim strPTCode As String = String.Empty, strPTName As String = String.Empty
        Dim strSUCCode As String = drData.Item("SUCCode").ToString
        Dim strSUCName As String = drData.Item("SUCName").ToString
        Dim strRcdRowId As String = drData.Item("RcdRowId").ToString

        Using dtPeriodCycle As DataTable = DAO.ExecQry(_DAL.GetPeriodCycle(drData.Item("Custid"), drData.Item("CitemCode"), drData.Item("RcdRowId")))
            If dtPeriodCycle.Rows.Count > 0 Then
                strCMCode = dtPeriodCycle.Rows(0).Item("CMCode").ToString
                strCMName = dtPeriodCycle.Rows(0).Item("CMName").ToString
                strPTCode = dtPeriodCycle.Rows(0).Item("PTCode").ToString
                strPTName = dtPeriodCycle.Rows(0).Item("PTName").ToString
            End If
        End Using
        If String.IsNullOrEmpty(strCMCode) Then
            Using dtCMCode As DataTable = DAO.ExecQry(_DAL.GetCMCode)
                For Each row As DataRow In dtCMCode.Rows
                    If row.IsNull("RefNo") = False Then
                        If row.Item("RefNo") = 1 Then
                            strCMCode = row.Item("CodeNo").ToString
                            strCMName = row.Item("Description").ToString
                            Exit For
                        End If
                    End If
                Next
            End Using
        End If
        If String.IsNullOrEmpty(strPTCode) Then
            Using dtPTCode As DataTable = DAO.ExecQry(_DAL.GetPTCode)
                For Each row As DataRow In dtPTCode.Rows
                    If row.Item("CodeNo") = 1 Then
                        strPTCode = row.Item("CodeNo").ToString
                        strPTName = row.Item("Description").ToString
                        Exit For
                    End If
                Next
            End Using
        End If
        '更新回SO033
        DAO.ExecNqry(_DAL.UpdRealCharge(strSUCCode, strSUCName, strCMCode, strCMName, strPTCode, strPTName, strRcdRowId))

        UCSO033 = True

        Return UCSO033
    End Function
    '收費結轉
    Public Function ChargeCutDate(ByVal dsCharge As DataSet) As CableSoft.BLL.Utility.RIAResult
        Dim strErrMsg As String = String.Empty
        Dim intSuccess As Integer = 0   '成功
        Dim intError As Integer = 0         '失敗
        Dim intErrCnt As Integer = 0        '異常
        Dim dtErrLog As DataTable = Nothing

        Dim trans As DbTransaction = Nothing
        Dim MyTrans As Boolean
        Dim cn As System.Data.Common.DbConnection = DAO.GetConn()

        Try
            Dim dtPara As DataTable = dsCharge.Tables("Para").Copy

            If DAO.Transaction Is Nothing Then
                MyTrans = True
                If cn.State = ConnectionState.Closed Then
                    cn.ConnectionString = Me.LoginInfo.ConnectionString
                    cn.Open()
                End If

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
            'CableSoft.BLL.Utility.Utility.SetClientInfo(DAO, LoginInfo.EntryId)
            If MyTrans Then CableSoft.BLL.Utility.Utility.SetClientInfo(DAO, LoginInfo.EntryId, CounterPayLanguage.ProcessKind4)

            '處理  intSuccess成功筆數  intError異常筆數  
            Dim result As RIAResult = So07xToSo033(dsCharge, intSuccess, intError, intErrCnt, strErrMsg, dtPara.Rows(0).Item("RealDate").ToString, dtErrLog, False, dtPara.Rows(0).Item("CutDay").ToString)
            If result.ResultBoolean = False Then
                If MyTrans Then trans.Rollback()
                Dim rData As DataSet = New DataSet()
                dtErrLog.TableName = "ErrData"
                rData.Merge(dtErrLog)
                'Return New RIAResult With {.ResultBoolean = False, .ErrorCode = -199, .ErrorMessage = CounterPayLanguage.SuccessMsg(intSuccess, intErrCnt), .ResultDataSet = rData}
                Return New RIAResult With {.ResultBoolean = False, .ErrorCode = -199, .ErrorMessage = CounterPayLanguage.SuccessMsg(0, intError), .ResultDataSet = rData}
            End If

            Dim strRtnMsg As String = CounterPayLanguage.SuccessMsg(intSuccess, intErrCnt)

            'Refresh
            '重取SO074A顯示到畫面
            Dim ds As DataSet = GetChargeTmp(dtPara)
            If dtErrLog IsNot Nothing Then
                dtErrLog.TableName = "ErrData"
                ds.Tables.Add(dtErrLog.Copy)
            End If

            If MyTrans Then
                trans.Commit()
                'cn.Close()
                'CableSoft.BLL.Utility.Utility.ClearClientInfo(DAO)
                'DAO.AutoCloseConn = True
                'DAO.Dispose()
                'cn.Dispose()
            End If

            Return New RIAResult With {.ResultBoolean = True, .ErrorCode = 0, .ErrorMessage = strRtnMsg, .ResultDataSet = ds.Copy}
        Catch ex As Exception
            If MyTrans Then trans.Rollback()
            Return New RIAResult With {.ErrorCode = -3, .ErrorMessage = "ChargeCutDate Error!!" & ex.ToString(), .ResultBoolean = False}
        Finally
            If MyTrans Then
                CableSoft.BLL.Utility.Utility.ClearClientInfo(DAO)
                'trans.Commit()
                cn.Close()
                DAO.AutoCloseConn = True
                DAO.Dispose()
                cn.Dispose()
            End If
        End Try
    End Function
    'SO074A更新到SO033入帳
    Public Function So07xToSo033(ByVal dsCharge As DataSet, ByRef intSuccess As Integer, ByRef intError As Integer, ByRef intErrCnt As Integer, ByRef strErrMsg As String,
                                ByVal strRealDate As String, ByRef dtErrLog As DataTable, ByVal strCutDay As String) As Boolean
        Dim result As RIAResult = So07xToSo033(dsCharge, intSuccess, intError, intErrCnt, strErrMsg, strRealDate, dtErrLog, False, strCutDay)
        Return result.ResultBoolean
    End Function
    'SO074A更新到SO033入帳
    Public Function So07xToSo033(ByVal dsCharge As DataSet, ByRef intSuccess As Integer, ByRef intError As Integer, ByRef intErrCnt As Integer, ByRef strErrMsg As String,
                                ByVal strRealDate As String, ByRef dtErrLog As DataTable, ByVal TranFlag As Boolean, ByVal strCutDay As String) As RIAResult
        Dim strUpdTime As String = CableSoft.BLL.Utility.DateTimeUtility.GetDTString(DateTime.Now)
        Dim NewUpdTime As DateTime = DateTime.Now
        Dim strWhere As String = String.Empty
        Dim blnUpdate As Boolean = True

        strWhere = String.Format("A.EntryEn ='{0}' And A.RealDate=To_Date('{1}','yyyy/MM/dd')", LoginInfo.EntryId, strRealDate)

        Using dtSO074A As DataTable = DAO.ExecQry(_DAL.GetChargeTmp2(LoginInfo.CompCode, strWhere))
            For Each dr074A As DataRow In dtSO074A.Rows
                blnUpdate = True
                Using dtSO033 As DataTable = DAO.ExecQry(_DAL.GetSimple2(), New Object() {LoginInfo.CompCode, dr074A.Item("BILLNO"), dr074A.Item("Item")})
                    If dtErrLog Is Nothing AndAlso dtSO033 IsNot Nothing Then
                        dtErrLog = dtSO033.Clone
                        dtErrLog.Columns.Add(New DataColumn With {.ColumnName = "CUSTNAME", .DataType = GetType(String)})
                        dtErrLog.AcceptChanges()
                    End If

                    If dtSO033 Is Nothing OrElse dtSO033.Rows.Count <= 0 Then
                        If dr074A.IsNull("RcdRowId") Then
                            dtSO033.Rows.Add(dtSO033.NewRow())
                            With dtSO033.Rows(0)
                                If .IsNull("RcdRowId") Then
                                    Dim AddrData As DataTable = DAO.ExecQry(_DAL.GetCustomerData(), New Object() {dr074A.Item("BillNo")})
                                    .Item("CUSTID") = dr074A.Item("CUSTID")
                                    .Item("BILLNO") = dr074A.Item("BILLNO")
                                    .Item("Item") = dr074A.Item("Item")
                                    .Item("OLDAMT") = dr074A.Item("ShouldAmt")
                                    .Item("OLDPERIOD") = dr074A.Item("RealPeriod")
                                    .Item("OLDSTARTDATE") = dr074A.Item("RealStartDate")
                                    .Item("OLDSTOPDATE") = dr074A.Item("RealStopDate")
                                    .Item("CREATETIME") = NewUpdTime
                                    .Item("CREATEEN") = dr074A.Item("EntryEn")
                                    .Item("COMPCODE") = dr074A.Item("CompCode")
                                    .Item("OLDCLCTEN") = dr074A.Item("CLCTEN")
                                    .Item("OLDCLCTNAME") = dr074A.Item("CLCTNAME")
                                    .Item("CUSTCOUNT") = 1
                                    .Item("REALAMT") = dr074A.Item("REALAMT")
                                    .Item("ClctAreaCode") = AddrData.Rows(0).Item("ClctAreaCode")
                                    .Item("ADDRNO") = AddrData.Rows(0).Item("AddrNo")
                                    .Item("STRTCODE") = AddrData.Rows(0).Item("STRTCODE")
                                    .Item("MDUID") = AddrData.Rows(0).Item("MduID")
                                    .Item("SERVCODE") = AddrData.Rows(0).Item("ServCode")
                                    .Item("CLASSCODE") = AddrData.Rows(0).Item("ClassCode")
                                    .Item("AreaCode") = AddrData.Rows(0).Item("AreaCode") '行政區代碼<1999/12/27>
                                    .Item("NodeNo") = AddrData.Rows(0).Item("NodeNo")
                                    .Item("FaciSeqNo") = dr074A.Item("FaciSeqNo")
                                    If .IsNull("FaciSeqNo") = False Then
                                        Dim FaciSNo As DataTable = DAO.ExecQry(_DAL.GetFaciSNo(), New Object() { .Item("FaciSeqNo")})
                                        If FaciSNo.Rows.Count > 0 Then
                                            .Item("FaciSNo") = FaciSNo.Rows(0).Item("FaciSNo")
                                        End If
                                    End If
                                    .Item("BillNoDate") = .Item("ShouldDate")
                                End If
                            End With
                        Else
                            '檢查是否已入帳或作廢
                            Dim drRow As DataRow = dtErrLog.NewRow
                            dtErrLog.Rows.Add(drRow)
                            CableSoft.BLL.Utility.Utility.CopyDataRow(dr074A, drRow)
                            drRow.Item("Note") = CounterPayLanguage.ChkNoData
                            drRow.Item("CUSTNAME") = dr074A.Item("CUSTNAME")
                            intErrCnt += 1
                            blnUpdate = False
                            Continue For
                        End If
                    End If

                    Dim strCustName As String = String.Empty
                    Dim str074aStartDate As String = String.Empty
                    Dim str033StartDate As String = String.Empty
                    If dr074A.IsNull("RealStartDate") = False Then str074aStartDate = dr074A.Item("RealStartDate").ToString
                    If dtSO033.Rows(0).IsNull("RealStartDate") = False Then str033StartDate = dtSO033.Rows(0).Item("RealStartDate").ToString
                    If dr074A.IsNull("CUSTNAME") = False Then strCustName = dr074A.Item("CUSTNAME").ToString

                    If dtSO033.Rows(0).Item("CancelFlag") = 1 Or ((dtSO033.Rows(0).IsNull("RealDate")) = False And dtSO033.Rows(0).IsNull("UCCode")) Then
                        '檢查是否已入帳或作廢
                        Dim drRow As DataRow = dtErrLog.NewRow
                        dtErrLog.Rows.Add(drRow)
                        CableSoft.BLL.Utility.Utility.CopyDataRow(dtSO033.Rows(0), drRow)
                        drRow.Item("Note") = CounterPayLanguage.ChargeClose
                        drRow.Item("CUSTNAME") = strCustName
                        intErrCnt += 1
                        blnUpdate = False
                        Continue For
                    Else
                        '起訖日被異動要log,仍要入帳
                        If str074aStartDate <> str033StartDate Then
                            Dim drRow As DataRow = dtErrLog.NewRow
                            dtErrLog.Rows.Add(drRow)
                            CableSoft.BLL.Utility.Utility.CopyDataRow(dtSO033.Rows(0), drRow)
                            drRow.Item("Note") = CounterPayLanguage.ChargeEdit
                            drRow.Item("CUSTNAME") = strCustName
                        End If

                        Using bll As New CableSoft.SO.BLL.Billing.Simple.Simple(LoginInfo, DAO)
                            With dtSO033.Rows(0)
                                CableSoft.BLL.Utility.Utility.CopyDataRow(dtSO033.Rows(0), dtSO033.Rows(0))
                                .Item("CitemCode") = dr074A.Item("CitemCode")
                                .Item("CitemName") = dr074A.Item("CitemName")
                                .Item("SHOULDDATE") = dr074A.Item("SHOULDDATE")
                                If .IsNull("RealDate") AndAlso .IsNull("FirstTime") Then
                                    .Item("FirstTime") = NewUpdTime
                                End If
                                'If strRealDate <> "" Then
                                '    .Item("REALDATE") = CDate(strRealDate)
                                'Else
                                '    .Item("REALDATE") = dr074A.Item("REALDATE")
                                'End If
                                If strCutDay <> "" Then
                                    .Item("REALDATE") = CDate(strCutDay)
                                Else
                                    .Item("REALDATE") = dr074A.Item("REALDATE")
                                End If

                                .Item("SHOULDAMT") = dr074A.Item("SHOULDAMT")
                                .Item("REALAMT") = dr074A.Item("REALAMT")
                                .Item("REALPERIOD") = dr074A.Item("REALPERIOD")

                                .Item("REALSTARTDATE") = dr074A.Item("REALSTARTDATE")
                                .Item("REALSTOPDATE") = dr074A.Item("REALSTOPDATE")
                                .Item("CLCTEN") = dr074A.Item("CLCTEN")
                                .Item("CLCTNAME") = dr074A.Item("CLCTNAME")
                                .Item("PTCODE") = dr074A.Item("PTCODE")
                                .Item("PTNAME") = dr074A.Item("PTNAME")
                                .Item("UPDTIME") = strUpdTime
                                .Item("UPDEN") = dr074A.Item("EntryEn")
                                .Item("NEWUPDTIME") = NewUpdTime
                                .Item("CMCODE") = dr074A.Item("CMCODE")
                                .Item("CMNAME") = dr074A.Item("CMNAME")
                                .Item("MANUALNO") = dr074A.Item("MANUALNO")
                                .Item("UCCode") = DBNull.Value
                                .Item("UCName") = DBNull.Value
                                .Item("STCODE") = dr074A.Item("STCODE")
                                .Item("STNAME") = dr074A.Item("STNAME")
                                .Item("Note") = dr074A.Item("Note")
                                .Item("ServiceType") = dr074A.Item("ServiceType")
                                .Item("CancelFlag") = dr074A.Item("CancelFlag")
                                .Item("CancelCode") = dr074A.Item("CancelCode")
                                .Item("CancelName") = dr074A.Item("CancelName")
                                'SO077 沒這些欄位 95/02/06 Jacky
                                .Item("BankCode") = dr074A.Item("BankCode")
                                .Item("BankName") = dr074A.Item("BankName")
                                .Item("AccountNo") = dr074A.Item("AccountNo")
                                .Item("AuthorizeNo") = dr074A.Item("AuthorizeNo")
                                .Item("AdjustFlag") = dr074A.Item("AdjustFlag")
                                .Item("NextPeriod") = dr074A.Item("NextPeriod")
                                .Item("NextAmt") = dr074A.Item("NextAmt")
                                .Item("InvSeqNo") = dr074A.Item("InvSeqNo")
                                '如實收日期由無值變為有值
                            End With
                            dtSO033.AcceptChanges()
                            dtSO033.TableName = "Simple"
                            '呼叫收費元件做入帳
                            If blnUpdate Then
                                If dtSO033.Columns.Contains("CTID") Then
                                    dtSO033.Columns.Remove("CTID")
                                End If
                                Dim result As RIAResult = bll.Save(EditMode.Edit, dtSO033.DataSet)
                                If result.ResultBoolean Then
                                    intSuccess += 1
                                Else
                                    strErrMsg = result.ErrorMessage
                                    intError += 1

                                    Dim drRow As DataRow = dtErrLog.NewRow
                                    dtErrLog.Rows.Add(drRow)
                                    CableSoft.BLL.Utility.Utility.CopyDataRow(dtSO033.Rows(0), drRow)
                                    drRow.Item("Note") = strErrMsg
                                    drRow.Item("CUSTNAME") = strCustName
                                End If
                            End If
                        End Using
                    End If
                End Using
            Next
        End Using
        If intError > 0 Then
            Return New RIAResult() With {.ResultBoolean = False}
        Else
            DAO.ExecNqry(_DAL.DeleteChargeTmp(strWhere))
            Return New RIAResult() With {.ResultBoolean = True}
        End If
    End Function
    '取得列印資料
    Public Function GetReportParams(dsConditions As DataSet) As DataSet 'Implements CableSoft.SL.Printing.Web.IReportData.GetReportParams

        Dim dtRpt As New DataTable
        dtRpt = dsConditions.Tables("ChargeTmp").Copy
        dtRpt.TableName = "Rpt"

        dsConditions.Tables.Add(dtRpt.Copy())

        Return dsConditions
    End Function

    ''' <summary>
    ''' 可否顯示
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CanView() As CableSoft.BLL.Utility.RIAResult
        Try
            Using soUtil As New CableSoft.SO.BLL.Utility.Utility(LoginInfo, DAO)
                Return soUtil.ChkPriv(LoginInfo.EntryId, "SO3318")
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    '修改付款種類
    Public Function EditPTData(ByVal dtPara As DataTable, ByVal PTCode As Integer, ByVal PTName As String, ByVal BillNo As String, ByVal Item As Integer) As CableSoft.BLL.Utility.RIAResult
        Dim strErrMsg As String = String.Empty

        Dim trans As DbTransaction = Nothing
        Dim MyTrans As Boolean
        Dim cn As System.Data.Common.DbConnection = DAO.GetConn()
        Try
            If DAO.Transaction Is Nothing Then
                MyTrans = True
                If cn.State = ConnectionState.Closed Then
                    cn.ConnectionString = Me.LoginInfo.ConnectionString
                    cn.Open()
                End If

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
            'CableSoft.BLL.Utility.Utility.SetClientInfo(DAO, LoginInfo.EntryId)
            If MyTrans Then CableSoft.BLL.Utility.Utility.SetClientInfo(DAO, LoginInfo.EntryId, CounterPayLanguage.ProcessKind3)

            DAO.ExecNqry(_DAL.UpdCharge("SO033", PTCode, PTName, BillNo, Item))
            DAO.ExecNqry(_DAL.UpdCharge("SO074A", PTCode, PTName, BillNo, Item))

            'Refresh
            '重取SO074A顯示到畫面
            Dim ds As DataSet = GetChargeTmp(dtPara)

            If MyTrans Then
                trans.Commit()
            End If

            Return New RIAResult With {.ResultBoolean = True, .ErrorCode = 0, .ErrorMessage = strErrMsg, .ResultDataSet = ds.Copy}
        Catch ex As Exception
            If MyTrans Then trans.Rollback()
            Return New RIAResult With {.ErrorCode = -2, .ErrorMessage = "EditPTData Error!!" & ex.ToString(), .ResultBoolean = False}
        Finally
            If MyTrans Then
                CableSoft.BLL.Utility.Utility.ClearClientInfo(DAO)
                'trans.Commit()
                cn.Close()
                DAO.AutoCloseConn = True
                DAO.Dispose()
                cn.Dispose()
            End If
        End Try
    End Function

    '修改客戶資訊
    Public Function SaveCustData(ByVal dtPara As DataTable, ByVal CarrierTypeCode As String, ByVal CarrierTypeName As String, ByVal CarrierId1 As String,
                                 ByVal LoveNum As String, ByVal CardLastNo As String, ByVal BillNo As String,
                                 ByVal MediaBillNo As String, ByVal dsInitData As DataSet) As CableSoft.BLL.Utility.RIAResult
        Dim trans As DbTransaction = Nothing
        Dim MyTrans As Boolean
        Dim cn As System.Data.Common.DbConnection = DAO.GetConn()
        Dim strErrMsg As String = String.Empty
        Dim result As RIAResult
        Dim intStartMobileAPI As Integer = 0
        Dim intStartLovenumAPI As Integer = 0
        '檢核IsDataOk
        Dim dtSystem As DataTable = dsInitData.Tables("System")
        If dtSystem.Rows.Count > 0 Then
            If dtSystem.Rows(0).IsNull("StartMobileAPI") = False Then intStartMobileAPI = dtSystem.Rows(0).Item("StartMobileAPI")
            If dtSystem.Rows(0).IsNull("StartLovenumAPI") = False Then intStartLovenumAPI = dtSystem.Rows(0).Item("StartLovenumAPI")
        End If
        Using objInvoice As New CableSoft.SO.BLL.Customer.Invoice.Invoice(LoginInfo, DAO)
            If intStartMobileAPI = 1 AndAlso String.IsNullOrEmpty(CarrierId1) = False Then
                'CD122參考號=1才需做驗證
                Dim dtCarrier As DataTable = DAO.ExecQry(_DAL.ChkCarrierType(CarrierTypeCode))
                If dtCarrier.Rows.Count > 0 Then
                    result = objInvoice.ChkCarrierId(CarrierId1, dtSystem)
                    If result.ResultBoolean = False Then
                        Return New RIAResult With {.ErrorCode = result.ErrorCode, .ErrorMessage = result.ErrorMessage, .ResultBoolean = False}
                    End If
                End If
            End If
            If intStartLovenumAPI = 1 AndAlso String.IsNullOrEmpty(LoveNum) = False Then
                result = objInvoice.ChkLoveNum(LoveNum, dtSystem)
                If result.ResultBoolean = False Then
                    Return New RIAResult With {.ErrorCode = result.ErrorCode, .ErrorMessage = result.ErrorMessage, .ResultBoolean = False}
                End If
            End If
        End Using

        Try
            If DAO.Transaction Is Nothing Then
                MyTrans = True
                If cn.State = ConnectionState.Closed Then
                    cn.ConnectionString = Me.LoginInfo.ConnectionString
                    cn.Open()
                End If

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
            'CableSoft.BLL.Utility.Utility.SetClientInfo(DAO, LoginInfo.EntryId)
            If MyTrans Then CableSoft.BLL.Utility.Utility.SetClientInfo(DAO, LoginInfo.EntryId, CounterPayLanguage.ProcessKind3)

            Dim strValue As String = "", strWhere As String = ""
            If String.IsNullOrEmpty(CarrierTypeCode) Then
                strValue = String.Format("{0},{1}", strValue, "CarrierTypeCode=null")
            Else
                strValue = String.Format("{0},{1}", strValue, "CarrierTypeCode='" & CarrierTypeCode & "'")
            End If
            If String.IsNullOrEmpty(CarrierId1) Then
                strValue = String.Format("{0},{1}", strValue, "CarrierId1=null")
            Else
                strValue = String.Format("{0},{1}", strValue, "CarrierId1='" & CarrierId1 & "'")
            End If

            If String.IsNullOrEmpty(LoveNum) Then
                strValue = String.Format("{0},{1}", strValue, "LoveNum=null")
            Else
                strValue = String.Format("{0},{1}", strValue, "LoveNum='" & LoveNum & "'")
            End If

            If String.IsNullOrEmpty(CardLastNo) Then
                strValue = String.Format("{0},{1}", strValue, "CardLastNo=null")
            Else
                strValue = String.Format("{0},{1}", strValue, "CardLastNo='" & CardLastNo & "'")
            End If

            If strValue.Substring(0, 1) = "," Then strValue = strValue.Substring(1)
            If String.IsNullOrEmpty(MediaBillNo) Then
                strWhere = " Where BillNo='" & BillNo & "'"
            Else
                strWhere = " Where MediaBillNo='" & MediaBillNo & "'"
            End If

            DAO.ExecNqry(_DAL.UpdCustData("SO033", strValue, strWhere))
            DAO.ExecNqry(_DAL.UpdCustData("SO074A", strValue, strWhere))

            'Refresh
            '重取SO074A顯示到畫面
            Dim ds As DataSet = GetChargeTmp(dtPara)

            If MyTrans Then
                trans.Commit()
            End If

            Return New RIAResult With {.ResultBoolean = True, .ErrorCode = 0, .ErrorMessage = strErrMsg, .ResultDataSet = ds.Copy}
        Catch ex As Exception
            If MyTrans Then trans.Rollback()
            Return New RIAResult With {.ErrorCode = -2, .ErrorMessage = "SaveCustData Error!!" & ex.ToString(), .ResultBoolean = False}
        Finally
            If MyTrans Then
                CableSoft.BLL.Utility.Utility.ClearClientInfo(DAO)
                'trans.Commit()
                cn.Close()
                DAO.AutoCloseConn = True
                DAO.Dispose()
                cn.Dispose()
            End If
        End Try
    End Function

    '信用卡刷卡  (判斷是否直接入帳(SO3318B不行,blnCredit=False))
    Public Function PaymentCharge(ByVal dsPara As DataSet, ByVal dsInitData As DataSet, ByVal blnCredit As Boolean) As CableSoft.BLL.Utility.RIAResult
        Dim trans As DbTransaction = Nothing
        Dim MyTrans As Boolean
        Dim cn As System.Data.Common.DbConnection = DAO.GetConn()
        Dim strErrMsg As String = String.Empty
        Dim result As RIAResult
        Dim CardBillNo As String = String.Empty
        Dim CardExpDate As String
        Dim blnDepositFlag As Boolean = False
        Dim intSTBCanOnlinePay As Integer = 0
        Dim blnDiffBill As Boolean = False
        Dim strValue As String = "", strValue2 As String = ""
        Dim strWhere As String = "", strCardLastNo As String = ""

        CardExpDate = dsPara.Tables(0).Rows(0).Item("CardExpDate").ToString
        CardExpDate = String.Format("{0}20{1}", CardExpDate.Substring(0, 2), CardExpDate.Substring(2, 2))

        blnDepositFlag = dsInitData.Tables("SystemCharge").Rows(0).Item("Para32") = 1
        intSTBCanOnlinePay = dsInitData.Tables("System").Rows(0).Item("STBCanOnlinePay")
        strCardLastNo = dsPara.Tables(0).Rows(0).Item("AccountNo").ToString.RightB(4)

        Try
            If DAO.Transaction Is Nothing Then
                MyTrans = True
                If cn.State = ConnectionState.Closed Then
                    cn.ConnectionString = Me.LoginInfo.ConnectionString
                    cn.Open()
                End If

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
            If MyTrans Then CableSoft.BLL.Utility.Utility.SetClientInfo(DAO, LoginInfo.EntryId, CounterPayLanguage.ProcessKind5)

            If dsInitData.Tables("Simple").Rows(0).IsNull("CardBillNo") Then
                Using bll As New CableSoft.SO.BLL.Utility.Utility(LoginInfo, DAO)
                    CardBillNo = bll.GetSequenceNo("S_SO033_CardBillNo", 8)
                    CardBillNo = "2" & CardBillNo.ToString.PadLeft(8, "0")
                End Using
                '信用卡單號回填SO033,SO074A
                strValue = "CardBillNo='" & CardBillNo & "',CardAccountNo='" & dsPara.Tables(0).Rows(0).Item("AccountNo") & "',AuthorizeNo='" & dsPara.Tables(0).Rows(0).Item("AuthorizeCode") & "',CardLastNo='" & strCardLastNo & "',CardCode=" & dsPara.Tables(0).Rows(0).Item("CardCode") & ",CardName='" & dsPara.Tables(0).Rows(0).Item("CardName") & "'"
                strValue2 = "CardBillNo='" & CardBillNo & "',AuthorizeNo='" & dsPara.Tables(0).Rows(0).Item("AuthorizeCode") & "',CardLastNo='" & strCardLastNo & "'"
                For Each drSimple As DataRow In dsInitData.Tables("Simple").Rows
                    strWhere = " Where BillNo='" & drSimple.Item("BillNo") & "' And Item=" & drSimple.Item("Item") & " And CancelFlag=0"
                    DAO.ExecNqry(_DAL.UpdCustData("SO033", strValue, strWhere))
                    DAO.ExecNqry(_DAL.UpdCustData("SO074A", strValue2, strWhere))
                Next
            Else
                CardBillNo = dsInitData.Tables("Simple").Rows(0).Item("CardBillNo").ToString
            End If

            If intSTBCanOnlinePay = 1 Then
                Dim dtDiffBill As DataTable = DAO.ExecQry(_DAL.GetDiffBill(CardBillNo))
                If dtDiffBill.Rows.Count > 0 Then blnDiffBill = True
            Else
                blnDiffBill = False
            End If

            Using objPay As New CableSoft.SO.BLL.Billing.Payment.Approve(LoginInfo, DAO) With {
                            .CardBillNo = CardBillNo,
                            .Credit = blnCredit,
                            .DepositFlag = blnDepositFlag,
                            .DiffBill = blnDiffBill,
                            .UpdTime = DateTime.Now,
                            .CardExpDate = CardExpDate,
                            .CardAccountNo = dsPara.Tables(0).Rows(0).Item("AccountNo"),
                            .AuthorizeCode = dsPara.Tables(0).Rows(0).Item("AuthorizeCode"),
                            .CardCode = dsPara.Tables(0).Rows(0).Item("CardCode"),
                            .RealDate = dsPara.Tables(0).Rows(0).Item("RealDate"),
                            .Amount = dsPara.Tables(0).Rows(0).Item("Amount"),
                            .CMCode = dsPara.Tables(0).Rows(0).Item("CMCode"),
                            .CMName = dsPara.Tables(0).Rows(0).Item("CMName"),
                            .PTCode = dsPara.Tables(0).Rows(0).Item("PTCode"),
                            .PTName = dsPara.Tables(0).Rows(0).Item("PTName"),
                            .ClctEn = dsPara.Tables(0).Rows(0).Item("ClctEn"),
                            .ClctName = dsPara.Tables(0).Rows(0).Item("ClctName")}

                result = objPay.Execute()
                If result.ResultBoolean = False Then
                    Return New RIAResult With {.ErrorCode = result.ErrorCode, .ErrorMessage = result.ErrorMessage, .ResultBoolean = False}
                End If
            End Using


            '重取SO074A顯示到畫面
            Dim ds As DataSet = GetChargeTmp(dsInitData.Tables("Para"))

            If MyTrans Then
                trans.Commit()
            End If

            Return New RIAResult With {.ResultBoolean = True, .ErrorCode = 0, .ErrorMessage = strErrMsg, .ResultDataSet = ds.Copy}
        Catch ex As Exception
            If MyTrans Then trans.Rollback()
            Return New RIAResult With {.ErrorCode = -2, .ErrorMessage = "PaymentCharge Error!!" & ex.ToString(), .ResultBoolean = False}
        Finally
            If MyTrans Then
                CableSoft.BLL.Utility.Utility.ClearClientInfo(DAO)
                'trans.Commit()
                cn.Close()
                DAO.AutoCloseConn = True
                DAO.Dispose()
                cn.Dispose()
            End If
        End Try
    End Function

    '檢核收費是否可以刪除登錄ChkChargeDel
    Public Function ChkChargeDel(ByVal dtCharge As DataTable) As CableSoft.BLL.Utility.RIAResult
        Dim strErrMsg As String = String.Empty
        Dim strWhere As String = String.Empty
        Dim dt As DataTable

        Try
            '970801 #4010 若該資料已無未收原因且已有實收日期者，請提示【已入實收或作廢】不可取消。
            If dtCharge.Rows.Count <= 0 Then
                Return New RIAResult With {.ResultBoolean = False, .ErrorCode = -1, .ErrorMessage = CounterPayLanguage.NoChargeData}
            Else
                For Each drSimple As DataRow In dtCharge.Rows
                    strWhere = " Where BillNo='" & drSimple.Item("BillNo") & "' And Item=" & drSimple.Item("Item") & " And RealDate is not null And UCCode is null"
                    dt = DAO.ExecQry(_DAL.ChkChargeData(strWhere))
                    If dt.Rows.Count > 0 Then
                        Return New RIAResult With {.ResultBoolean = False, .ErrorCode = -3, .ErrorMessage = CounterPayLanguage.NoDelCharge}
                    End If
                Next
            End If

            Return New RIAResult With {.ResultBoolean = True, .ErrorCode = 0, .ErrorMessage = strErrMsg}
        Catch ex As Exception
            Return New RIAResult With {.ResultBoolean = False, .ErrorCode = -2, .ErrorMessage = "ChkChargeDel Error!!" & ex.ToString()}
        End Try
    End Function

    '信用卡退刷  (判斷是否直接入帳(SO3318B不行,blnCredit=False))
    Public Function PaymentChargeDel(ByVal dsInitData As DataSet, ByVal blnCredit As Boolean) As CableSoft.BLL.Utility.RIAResult
        Dim trans As DbTransaction = Nothing
        Dim MyTrans As Boolean
        Dim cn As System.Data.Common.DbConnection = DAO.GetConn()
        Dim strErrMsg As String = String.Empty
        Dim result As RIAResult
        Dim CardBillNo As String = String.Empty
        Dim blnDepositFlag As Boolean = False
        Dim intSTBCanOnlinePay As Integer = 0
        Dim blnDiffBill As Boolean = False
        Dim strValue As String = "", strWhere As String = ""

        blnDepositFlag = dsInitData.Tables("SystemCharge").Rows(0).Item("Para32") = 1
        intSTBCanOnlinePay = dsInitData.Tables("System").Rows(0).Item("STBCanOnlinePay")

        Try
            If DAO.Transaction Is Nothing Then
                MyTrans = True
                If cn.State = ConnectionState.Closed Then
                    cn.ConnectionString = Me.LoginInfo.ConnectionString
                    cn.Open()
                End If

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
            If MyTrans Then CableSoft.BLL.Utility.Utility.SetClientInfo(DAO, LoginInfo.EntryId, CounterPayLanguage.ProcessKind5)

            If dsInitData.Tables("Simple").Rows(0).IsNull("CardBillNo") = False Then
                CardBillNo = dsInitData.Tables("Simple").Rows(0).Item("CardBillNo").ToString
            End If

            If intSTBCanOnlinePay = 1 Then
                Dim dtDiffBill As DataTable = DAO.ExecQry(_DAL.GetDiffBill(CardBillNo))
                If dtDiffBill.Rows.Count > 0 Then blnDiffBill = True
            Else
                blnDiffBill = False
            End If

            Using objPay As New CableSoft.SO.BLL.Billing.Payment.DepositReversal(LoginInfo, DAO) With {
                            .CardBillNo = CardBillNo,
                            .Credit = blnCredit,
                            .DepositFlag = blnDepositFlag,
                            .DiffBill = blnDiffBill,
                            .UpdTime = DateTime.Now}

                result = objPay.Execute()
                If result.ResultBoolean = False Then
                    Return New RIAResult With {.ErrorCode = result.ErrorCode, .ErrorMessage = result.ErrorMessage, .ResultBoolean = False}
                End If
            End Using

            If dsInitData.Tables("Simple").Rows.Count > 0 Then
                '清除SO033信用卡單號
                strValue = "CardBillNo=NULL"
                For Each drSimple As DataRow In dsInitData.Tables("Simple").Rows
                    strWhere = " Where BillNo='" & drSimple.Item("BillNo") & "' And Item=" & drSimple.Item("Item") & " And CancelFlag=0"
                    DAO.ExecNqry(_DAL.UpdCustData("SO033", strValue, strWhere))
                    'DAO.ExecNqry(_DAL.UpdCustData("SO074A", strValue, strWhere))
                Next
            End If

            '重取SO074A顯示到畫面
            Dim ds As DataSet = GetChargeTmp(dsInitData.Tables("Para"))

            If MyTrans Then
                trans.Commit()
            End If

            Return New RIAResult With {.ResultBoolean = True, .ErrorCode = 0, .ErrorMessage = strErrMsg, .ResultDataSet = ds.Copy}
        Catch ex As Exception
            If MyTrans Then trans.Rollback()
            Return New RIAResult With {.ErrorCode = -2, .ErrorMessage = "PaymentChargeDel Error!!" & ex.ToString(), .ResultBoolean = False}
        Finally
            If MyTrans Then
                CableSoft.BLL.Utility.Utility.ClearClientInfo(DAO)
                'trans.Commit()
                cn.Close()
                DAO.AutoCloseConn = True
                DAO.Dispose()
                cn.Dispose()
            End If
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
