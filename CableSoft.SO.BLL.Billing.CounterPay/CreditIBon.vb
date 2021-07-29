Imports System.Data
Imports System.Data.Common
Imports CableSoft.BLL.Utility
Imports CableSoft.SO.BLL.Utility
Imports util = CableSoft.BLL.Utility
Imports CableSoft.Utility.DataAccess
Imports System.IO
Imports System.Threading
'Imports CableSoft.SO.BLL.Billing.PayCommand
Imports CreditIBonLanguage = CableSoft.BLL.Language.SO61.CreditIBonLanguage

Public Class CreditIBon
    Inherits CableSoft.BLL.Utility.BLLBasic
    Implements IDisposable

    Private _DAL As New CreditIBonDALMultiDB(Me.LoginInfo.Provider)
    Private ThreadCount As Integer = 20

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

    '前端呼叫資料查詢
    Public Function GetDetailData(ByVal strSQL As String) As DataTable
        For Each pInfo As Reflection.PropertyInfo In LoginInfo.GetType.GetProperties
            Dim FieldName As String = "LoginInfo." & pInfo.Name
            If strSQL.ToUpper.IndexOf(String.Format("[{0}]", FieldName.ToUpper)) >= 0 Then
                Dim FieldValue As String = pInfo.GetValue(LoginInfo, Nothing)
                strSQL = CableSoft.BLL.Utility.Utility.GetFieldValueSQL(_DAL.Sign, strSQL, FieldName, FieldValue, Nothing)
            End If
        Next
        Dim dt As DataTable = DAO.ExecQry(strSQL.ToUpper)
        Return dt
    End Function

    'MSSQL的連線資訊,Oracle的連線資訊(公司別,連線字串),已撈取需處理的資料,參數
    Public Function Execute(dtReadConn As DataTable, dtWriteConn As DataTable, ChargeData As DataTable, paraObj As Object) As Boolean
        Dim GroupRows = ChargeData.AsEnumerable.GroupBy(Function(list) list.Item("SHOPID"))
        Dim myThread(ThreadCount - 1) As Thread
        ThreadCount = GroupRows.Count

        For Each GroupRow In GroupRows
            Dim cThread As Thread = Nothing
            Do While cThread Is Nothing
                cThread = GetCanUseThread(myThread, paraObj, New ParameterizedThreadStart(AddressOf ExecuteGo))
                Thread.Sleep(10)
            Loop
            ChildThreadGo(LoginInfo, GroupRow, cThread, paraObj, dtReadConn, dtWriteConn)
            cThread.Join()
        Next
        ChkThreadFinish(myThread)
        Return True
    End Function
    '檢查執行緒是否已執行結束
    Private Function ChkThreadFinish(myThread As Thread()) As Boolean
        Do While True
            Dim ChkOk As Boolean = True
            For Each TempThread As Thread In myThread
                If Not (TempThread Is Nothing OrElse TempThread.IsAlive = False) Then
                    ChkOk = False
                    Exit For
                End If
            Next
            If ChkOk Then
                Exit Do
            End If
            Thread.Sleep(10)
        Loop
        Return True
    End Function
    '呼叫子執行緒
    Private Sub ChildThreadGo(LoginInfo As CableSoft.BLL.Utility.LoginInfo,
                            ChooseRows As IGrouping(Of Object, DataRow),
                            RealThread As Thread, paraObj As Object,
                            dtReadConn As DataTable, dtWriteConn As DataTable)
        Dim strConnStringR As String = String.Empty 'MSSQL的連線字串
        Dim strConnStringW As String = String.Empty 'Oracle的連線字串
        Dim strCompCodeR As String = paraObj(0) 'MSSQL的公司別
        Dim strCompCodeW As String = paraObj(0) 'Oracle的公司別
        Dim ReadCompStr As String = paraObj(1)         'MSSQL的Provider
        Dim WriteCompStr As String = paraObj(2)         'Oracle的Provider
        Dim dtMSSetting As DataTable = paraObj(3)    '設定Table中的公司資訊dtMSSetting
        Dim strPath As String = paraObj(4)                      '程式執行路徑
        Dim NewLoginInfo As New CableSoft.BLL.Utility.LoginInfo  'MSSQL用的
        Dim NewLoginInfo2 As New CableSoft.BLL.Utility.LoginInfo  'ORACLE用的
        Dim RLoginInfo As CableSoft.BLL.Utility.LoginInfo = CableSoft.BLL.Utility.Utility.ConvertTo(LoginInfo, NewLoginInfo)
        Dim WLoginInfo As CableSoft.BLL.Utility.LoginInfo = CableSoft.BLL.Utility.Utility.ConvertTo(LoginInfo, NewLoginInfo2)

        '取得公司別所對應的連線字串
        If Not ChooseRows Is Nothing Then
            If ChooseRows.Count > 0 Then
                'If Not String.IsNullOrEmpty(strCompCodeW) Then
                '判斷此資料群組的公司別是否有設定在Oracle的公司別中,有設定才可做
                'If (String.Format(",{0},", WriteCompStr)).Contains(String.Format(",{0},", strCompCodeW)) Then
                '取得Config MSSQL Table中的公司資訊
                For Each dr As DataRow In dtReadConn.AsEnumerable.Where(Function(list) list.Item("CompCode") = strCompCodeR)
                    strConnStringR = dr.Item("Connection").ToString
                    Exit For
                Next
                RLoginInfo.CompCode = Integer.Parse(strCompCodeR)
                RLoginInfo.ConnectionString = strConnStringR

                '取得Config ORACLE Table中的公司資訊
                For Each dr As DataRow In dtWriteConn.AsEnumerable.Where(Function(list) list.Item("CompCode") = strCompCodeW)
                    strConnStringW = dr.Item("Connection").ToString
                    Exit For
                Next
                WLoginInfo.CompCode = Integer.Parse(strCompCodeW)
                WLoginInfo.ConnectionString = strConnStringW
                WLoginInfo.Provider = WriteCompStr

                '斷線重連
                If ChkConOK(strConnStringR, strPath, Integer.Parse(strCompCodeR), RLoginInfo) Then
                    If ChkConOK(strConnStringW, strPath, Integer.Parse(strCompCodeW), WLoginInfo) Then
                        Dim Thread As Thread = RealThread
                        '有資料則要建立執行緒
                        paraObj(5) = RLoginInfo
                        paraObj(6) = WLoginInfo
                        paraObj(7) = ChooseRows
                        '3.啟動執行緒
                        Thread.Start(paraObj)
                    End If
                End If
                'End If
                'End If
            End If
        End If
    End Sub
    '取得可使用的執行緒
    Private Function GetCanUseThread(myThreads() As Thread, paraObj As Object, ExecuteSub As ParameterizedThreadStart) As Thread
        Dim TempThread As Thread = Nothing
        Try
            For intLoop As Integer = 0 To ThreadCount - 1
                TempThread = myThreads(intLoop)
                If TempThread Is Nothing OrElse TempThread.IsAlive = False Then
                    '2.建立Thread 類別
                    Dim myPar As ParameterizedThreadStart = ExecuteSub
                    'Dim myRun As ThreadStart = New ThreadStart(AddressOf RunSample01)
                    If TempThread IsNot Nothing Then
                        TempThread = Nothing
                    End If
                    TempThread = New Thread(myPar)
                    myThreads(intLoop) = TempThread
                    Exit For
                Else
                    TempThread = Nothing
                End If
            Next
        Catch ex As Exception
            WriteErrorLog(paraObj(7), LoginInfo.CompCode, ex.ToString, "GetCanUseThread")
        End Try
        Return TempThread
    End Function
    '做資料回填,呼叫命令元件
    Private Sub ExecuteGo(ParaObj As Object)
        '0.MSSQL跟ORACLE要設定相同對應的公司別
        '1.MSSQL的Provider
        '2.Oracle的Provider
        '3.設定Table中的公司資訊dtMSSetting
        '4.程式執行路徑
        '5.後端MSSQL的RLoginInfo
        '6.後端ORACLE的WLoginInfo
        '7.依公司別分組的資料Rows

        Dim WriteCompStr As String = ParaObj(0)     'MSSQL跟ORACLE要設定相同對應的公司別
        'Dim RProvider As String = ParaObj(1)            'MSSQL的Provider
        'Dim WProvider As Integer = ParaObj(2)    'Oracle的Provider
        Dim dtMSSetting As DataTable = ParaObj(3)      '設定Table中的公司資訊dtMSSetting
        Dim strPath As String = ParaObj(4)                          '程式執行路徑
        Dim RLoginInfo As LoginInfo = ParaObj(5)
        Dim WLoginInfo As LoginInfo = ParaObj(6)
        Dim ChooseRows As IGrouping(Of Object, DataRow) = ParaObj(7)

        Dim tranR As DbTransaction = Nothing
        Dim blnTransR As Boolean = False
        Dim tranW As DbTransaction = Nothing
        Dim blnTransW As Boolean = False
        Dim intCustid As Integer = 0
        Dim strMediaBillNo As String = String.Empty
        Dim lngAmount As Long = 0
        Dim dtSystem As DataTable = Nothing
        Dim intCrossCustCombine As Integer = 0
        Dim strErrMsg As String = String.Empty
        Dim strProcessTime As String = String.Empty

        Try
            Using childWriteDao As New DAO(WLoginInfo.Provider, WLoginInfo.ConnectionString)
                Dim connW As DbConnection = childWriteDao.GetConn()
                childWriteDao.AutoCloseConn = False
                connW.ConnectionString = WLoginInfo.ConnectionString
                connW.Open()

                'SO041
                Using Utility As New CableSoft.SO.BLL.Utility.Utility(WLoginInfo, childWriteDao)
                    dtSystem = Utility.GetSystem(BLL.Utility.SystemTableType.System, "*", Nothing)
                    If dtSystem.Rows(0).IsNull("CrossCustCombine") = False Then
                        intCrossCustCombine = Integer.Parse(dtSystem.Rows(0).Item("CrossCustCombine"))
                    End If
                End Using

                Dim strShopID As String = String.Empty
                Dim strTableName As String = String.Empty

                For Each dr As DataRow In dtMSSetting.Rows
                    If dr.Item(0) = RLoginInfo.CompCode Then
                        If dr.IsNull(1) = False Then strShopID = dr.Item(1).ToString
                        If dr.IsNull(2) = False Then strTableName = dr.Item(2).ToString
                        Exit For
                    End If
                Next

                For Each dr As DataRow In ChooseRows
                    Try
                        intCustid = dr.Item("Custid")
                        strMediaBillNo = dr.Item("MediaBillNo")
                        lngAmount = dr.Item("Amount")
                        strErrMsg = ""
                        'strProcessTime = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")
                        strProcessTime = String.Format("{0:yyyy/MM/dd HH:mm:ss}", DateTime.Now)

                        If dr.Item("Status") = CreditIBonLanguage.fStatusNot Then      '狀態為未處理才做
                            '建立讀取資料的DAO
                            Using childReadDao As New DAO(RLoginInfo.Provider, RLoginInfo.ConnectionString)
                                Dim connR As DbConnection = childReadDao.GetConn()
                                childReadDao.AutoCloseConn = False
                                connR.ConnectionString = RLoginInfo.ConnectionString
                                connR.Open()

                                '取SO033完整資料
                                Using dtCmd As DataTable = childWriteDao.ExecQry(_DAL.GetRealCharge(intCustid, strMediaBillNo, intCrossCustCombine))
                                    If dtCmd.Rows.Count > 0 Then
                                        Dim Parameters As Object() = Nothing

                                        tranR = connR.BeginTransaction
                                        childReadDao.AutoCloseConn = False
                                        blnTransR = True
                                        childReadDao.Transaction = tranR

                                        tranW = connW.BeginTransaction
                                        childWriteDao.AutoCloseConn = False
                                        blnTransW = True
                                        childWriteDao.Transaction = tranW

                                        'Dim strServicetype As String = String.Empty
                                        Dim strUCCode As String = String.Empty
                                        Dim strUCName As String = String.Empty
                                        'strServicetype = dtCmd.Rows(0).Item("ServiceType").ToString
                                        '取得對應已收UCCODE
                                        Using dtUCCode As DataTable = childWriteDao.ExecQry(_DAL.GetUCCode())
                                            If dtUCCode.Rows.Count > 0 Then
                                                strUCCode = dtUCCode.Rows(0).Item("CodeNo")
                                                strUCName = dtUCCode.Rows(0).Item("Description")
                                            End If
                                        End Using

                                        '檢核收費資料是否可執行入帳
                                        If Not IsDataOK(childWriteDao, WLoginInfo, strPath, intCustid, strMediaBillNo, lngAmount, intCrossCustCombine, strErrMsg) Then
                                            WriteErrorLog(strPath, WLoginInfo.CompCode, strErrMsg, "IsDataOK")
                                            'childReadDao.ExecNqry(_DAL.UpdMSData(strShopID, strTableName, strErrMsg, strMediaBillNo))
                                            'childReadDao.ExecNqry(_DAL.UpdMSData(), New Object() {strShopID, strTableName, CreditIBonLanguage.fStatusYes, strMediaBillNo, strProcessTime})
                                            'childReadDao.ExecNqry(_DAL.UpdMSData(), New Object() {strShopID, strTableName, CreditIBonLanguage.fStatusYes, strMediaBillNo, DateTime.Now})
                                            childReadDao.ExecNqry(_DAL.UpdMSData(strShopID, strTableName, CreditIBonLanguage.fStatusYes, strMediaBillNo), New Object() {DateTime.Now})
                                            If blnTransR Then tranR.Commit()
                                            If blnTransW Then tranW.Rollback()
                                            blnTransR = False
                                            blnTransW = False
                                        Else
                                            Dim result As RIAResult = Nothing
                                            Dim DateTimeUtility As New CableSoft.BLL.Utility.DateTimeUtility()
                                            Dim strUpdTime As String = DateTimeUtility.GetDTString(DateTime.Now)
                                            Dim dtCharge As DataTable = Nothing
                                            Try
                                                '送命令前 SO033 異動UCCODE
                                                If String.IsNullOrEmpty(strUCCode) = False Then
                                                    childWriteDao.ExecNqry(_DAL.UpdRealCharge(strUCCode, strUCName, WLoginInfo.EntryName, strUpdTime, strMediaBillNo))
                                                End If
                                                '建立Corey 送命令處理元件  'dll=CableSoft.SO.BLL.Billing.PayCommand

                                                Using SOUtil As New CableSoft.SO.BLL.Billing.PayCommand.PayCommand(WLoginInfo, childWriteDao)
                                                    'Public Function Execute(dtMediaBillNo As System.Data.DataTable, intSendDelayRecTime As Integer) As RIAResult
                                                    dtCharge = childWriteDao.ExecQry(_DAL.GetRealCharge(intCustid, strMediaBillNo, intCrossCustCombine))
                                                    result = SOUtil.Execute(dtCharge, 3)
                                                End Using
                                                '呼叫PayCommand資料回寫OK
                                                If result.ResultBoolean Then
                                                    '元件命令執行成功後,要更新VOD點數
                                                    Using vodCredit As New CableSoft.SO.BLL.Billing.Utility.Utility(WLoginInfo, childWriteDao)
                                                        For Each row As DataRow In dtCharge.Rows
                                                            vodCredit.UpdateVODPoint(row)
                                                        Next
                                                    End Using
                                                    'childReadDao.ExecNqry(_DAL.UpdMSData(), New Object() {strShopID, strTableName, CreditIBonLanguage.fStatusYes, strMediaBillNo, strProcessTime})
                                                    'childReadDao.ExecNqry(_DAL.UpdMSData(), New Object() {strShopID, strTableName, CreditIBonLanguage.fStatusYes, strMediaBillNo, DateTime.Now})
                                                    childReadDao.ExecNqry(_DAL.UpdMSData(strShopID, strTableName, CreditIBonLanguage.fStatusYes, strMediaBillNo), New Object() {DateTime.Now})
                                                    If blnTransR Then tranR.Commit()
                                                    If blnTransW Then tranW.Commit()
                                                    blnTransR = False
                                                    blnTransW = False
                                                Else
                                                    WriteErrorLog(strPath, WLoginInfo.CompCode, result.ErrorMessage, "ExecuteGo-PayCommand-ERR1")
                                                    'childReadDao.ExecNqry(_DAL.UpdMSData(strShopID, strTableName, CreditIBonLanguage.fCmdError, strMediaBillNo))
                                                    'childReadDao.ExecNqry(_DAL.UpdMSData(), New Object() {strShopID, strTableName, CreditIBonLanguage.fStatusYes, strMediaBillNo, strProcessTime})
                                                    'childReadDao.ExecNqry(_DAL.UpdMSData(), New Object() {strShopID, strTableName, CreditIBonLanguage.fStatusYes, strMediaBillNo, DateTime.Now})
                                                    childReadDao.ExecNqry(_DAL.UpdMSData(strShopID, strTableName, CreditIBonLanguage.fStatusYes, strMediaBillNo), New Object() {DateTime.Now})
                                                    If blnTransR Then tranR.Commit()
                                                    If blnTransW Then tranW.Rollback()
                                                    blnTransR = False
                                                    blnTransW = False
                                                End If
                                            Catch ex As Exception
                                                WriteErrorLog(strPath, WLoginInfo.CompCode, ex.ToString, "ExecuteGo-PayCommand-ERR2")
                                            Finally
                                                If blnTransR Then tranR.Rollback()
                                                If blnTransW Then tranW.Rollback()
                                                blnTransR = False
                                                blnTransW = False
                                            End Try
                                        End If
                                    Else    'SO033取不到收費資料 log err
                                        WriteErrorLog(strPath, WLoginInfo.CompCode, CreditIBonLanguage.fNoCharge, "ExecuteGo-GetRealCharge-ERR")
                                        'childReadDao.ExecNqry(_DAL.UpdMSData(strShopID, strTableName, CreditIBonLanguage.fNoCharge, strMediaBillNo))
                                        'childReadDao.ExecNqry(_DAL.UpdMSData(), New Object() {strShopID, strTableName, CreditIBonLanguage.fStatusYes, strMediaBillNo, strProcessTime})
                                        childReadDao.ExecNqry(_DAL.UpdMSData(strShopID, strTableName, CreditIBonLanguage.fStatusYes, strMediaBillNo), New Object() {DateTime.Now})
                                        If blnTransR Then tranR.Commit()
                                        If blnTransW Then tranW.Rollback()
                                        blnTransR = False
                                        blnTransW = False
                                    End If
                                End Using
                                connR.Close()
                                connR.Dispose()
                            End Using
                        End If
                    Catch ex As Exception
                        WriteErrorLog(strPath, WLoginInfo.CompCode, ex.ToString, "ExecuteGo-DLLErr")
                    End Try
                Next
                connW.Close()
                connW.Dispose()
            End Using
        Catch ex As Exception
            WriteErrorLog(strPath, WLoginInfo.CompCode, ex.ToString, "ExecuteGo")
        End Try
    End Sub
    '檢查連線狀態
    Private Function ChkConOK(ByVal connstring As String, ByVal strPath As String, CompCode As Integer, XLoginInfo As LoginInfo) As Boolean
        ChkConOK = False
        Dim ErrMsg As String = String.Empty

        Try
            If String.IsNullOrEmpty(connstring) Then
                WriteErrorLog(strPath, CompCode, CreditIBonLanguage.NoConnStr, "ChkConOK")
                ChkConOK = False
            Else

                Using checkDao As New DAO(XLoginInfo.Provider, connstring)
                    Using cn As DbConnection = checkDao.GetConn()
                        For intLoop As Integer = 0 To 4
                            If cn.State = ConnectionState.Closed Then cn.Open()
                            If cn.State = ConnectionState.Open Then
                                ChkConOK = True
                                Exit For
                            Else
                                ErrMsg = CreditIBonLanguage.NotConnMsg(intLoop + 1)
                                WriteErrorLog(strPath, CompCode, ErrMsg, "ChkConOK")
                            End If
                        Next
                        cn.Close()
                    End Using
                    'ChkConOK = True
                End Using
            End If
            Return ChkConOK
        Catch ex As Exception
            ErrMsg = CreditIBonLanguage.NotConnStrMsg(ex.ToString, connstring)
            WriteErrorLog(strPath, CompCode, ErrMsg, "ChkConOK")
        End Try
    End Function
    '發生錯誤寫入ErrLog
    Private Function WriteErrorLog(strPath As String, CompCode As Integer, ErrorMsg As String, ModuleName As String) As Boolean
        Try
            SyncLock Me
                Using myWriter As New System.IO.StreamWriter(strPath & "\CredidIBONErr.txt", True)
                    myWriter.WriteLine(String.Format("發生時間: {0} ,ModuleName: {1} ,第 {2} 家公司別 ,錯誤訊息: {3}", Format(DateTime.Now, "yyyy/MM/dd HH:mm:ss"), ModuleName, CompCode, ErrorMsg))
                    myWriter.Close()
                End Using
            End SyncLock
        Catch ex As Exception
        End Try
        Return True
    End Function
    'IsDataOK
    Private Function IsDataOK(chkDao As DAO, XLoginInfo As LoginInfo, ByVal strPath As String, ByVal intCustid As Integer, ByVal strMediaBillNo As String,
                                                    ByVal lngAmount As Long, ByVal intCrossCustCombine As Integer, ByRef strErrMsg As String) As Boolean
        IsDataOK = False
        Try
            '判斷SO033是否有該媒体單號
            Dim dt As DataTable = Nothing, dt2 As DataTable = Nothing
            Dim intCount As Integer = 0, intCount2 As Integer = 0

            dt = chkDao.ExecQry(_DAL.GetChargeCnt(strMediaBillNo))
            If dt.Rows.Count > 0 Then
                If dt.Rows(0)(0) = 0 Then
                    strErrMsg = CreditIBonLanguage.fNotSMS
                    Return False
                End If
            Else
                strErrMsg = CreditIBonLanguage.fNotSMS
                Return False
            End If
            '如果有媒体單號再判斷媒体單號是否屬於該客編
            dt = chkDao.ExecQry(_DAL.GetChargeCustid(strMediaBillNo, intCustid))
            If dt.Rows.Count > 0 Then
                If dt.Rows(0)(0) > 0 Then
                    strErrMsg = CreditIBonLanguage.fNotSMS
                    Return False
                End If
            End If
            '是否金額不符   '990127 #5499 調整金額加總時要過濾作廢
            dt = chkDao.ExecQry(_DAL.GetChargeAmount(strMediaBillNo, intCustid, intCrossCustCombine))
            If dt.Rows.Count > 0 Then
                If Long.Parse(dt.Rows(0)(0)) <> lngAmount Then
                    strErrMsg = CreditIBonLanguage.fNotAmount
                    Return False
                End If
            End If
            '是否作廢   '990127 #5499 調整若整張全都作廢才能算作廢,單一筆不算(總筆數=總作廢筆數)
            dt = chkDao.ExecQry(_DAL.GetChargeCancel(strMediaBillNo, intCustid, intCrossCustCombine))
            dt2 = chkDao.ExecQry(_DAL.GetChargeCancel2(strMediaBillNo, intCustid, intCrossCustCombine))
            If dt.Rows.Count > 0 Then
                intCount = dt.Rows(0)(0)
            End If
            If dt2.Rows.Count > 0 Then
                intCount2 = dt2.Rows(0)(0)
            End If
            If intCount = intCount2 Then
                strErrMsg = CreditIBonLanguage.fCancelFlag
                Return False
            End If
            '是否已收   '990224 #5499 測試報告,因作廢會填RealDate,要增加條件過濾  '990503 #5641 調整判斷已收的條件 '990519 #5564 調整規格,已收的判斷,CD013.RefNo=3,7 or PayOk=1
            dt = chkDao.ExecQry(_DAL.GetChargePayOK(strMediaBillNo, intCustid, intCrossCustCombine))
            If dt.Rows.Count > 0 Then
                If dt.Rows(0)(0) > 0 Then
                    strErrMsg = CreditIBonLanguage.fUCCode
                    Return False
                End If
            End If

            IsDataOK = True
        Catch ex As Exception
            WriteErrorLog(strPath, XLoginInfo.CompCode, ex.ToString, "IsDataOK")
        End Try
        Return IsDataOK
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
