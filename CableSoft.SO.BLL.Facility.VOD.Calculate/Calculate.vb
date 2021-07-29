Imports System.Data.Common
Imports CableSoft.BLL.Utility

Public Class Calculate
    Inherits BLLBasic
    Implements IDisposable
    Private _DAL As New CalculateDAL(Me.LoginInfo.Provider)
    
    Private Const tbCloseParaName As String = "ClosePara"
    Private Const tbChargeParaName As String = "ChargePara"
    Private Const tbDefaultCitemName As String = "DefaultCitem"
    Private Const tbDefaultCMCodeName As String = "DefaultCMCode"
    Private Const tbDefaultUCCodeName As String = "DefaultUCCode"
    Private Const tbDefaultSalePointCodeName As String = "DefaultSalePointCode"
    Private Const tbServiceTypeName As String = "ServiceType"
    Private Const tbCompCodeName As String = "CompCode"
    Private Const ParaTableName As String = "VODCalculate"
    Private Const tbChooseName As String = "Choose"
    Private Const tbPrivName As String = "Priv"

    Private strViewXLSName As String = Nothing
    Private Const tbExcelName As String = "ImportExcel"
    Private Const ScreenParaTableName As String = "Para"
    Private Lang As New CableSoft.BLL.Language.SO61.CalculateLanguage()
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
    Public Function DeleteView(ByVal ViewName As String) As RIAResult
        Dim result As RIAResult = New RIAResult With {.ResultBoolean = True, .ErrorCode = 0}
        Try
            For i As Integer = 0 To ViewName.Split(",")(0).Split(";").Count - 1
                If i = 0 Then
                    DAO.ExecNqry(_DAL.DeleteView(ViewName.Split(",")(0).Split(";")(0)))
                Else
                    DAO.ExecNqry(_DAL.DeleteTmpTable(ViewName.Split(",")(0).Split(";")(1)))
                End If
            Next
           
        Catch ex As Exception
            result.ResultBoolean = False
            result.ErrorCode = -1
            result.ErrorMessage = ex.ToString
        End Try
        
        Return result
    End Function
    Public Function ChkAuthority(ByVal Mid As String) As RIAResult
        Dim result As New RIAResult() With {.ErrorCode = 0, .ErrorMessage = Nothing, .ResultBoolean = True}
        Try
            Using obj As New CableSoft.SO.BLL.Utility.Utility(Me.LoginInfo, DAO)
                result = obj.ChkPriv(LoginInfo.EntryId, Mid)
                obj.Dispose()
            End Using
            If Not result.ResultBoolean Then
                result.ErrorCode = -1        
            End If
        
        Catch ex As Exception
            result.ErrorMessage = ex.ToString
            result.ResultBoolean = False
            result.ErrorCode = -2
        Finally

        End Try
        Return result

    End Function
    Public Function CalculateBill(ByVal dsSource As DataSet) As RIAResult
        Dim dsResult As New DataSet
        Dim result As RIAResult = New RIAResult With {.ResultBoolean = True, .ErrorCode = 0}
        Dim ResultViewName As String = Nothing
        Dim isVodData As Integer = 0
        Dim noVodData As Integer = 0
        If (dsSource.Tables.Contains(tbExcelName)) AndAlso
            (dsSource.Tables(tbExcelName).Rows.Count > 0) Then
            strViewXLSName = "TMP_" & DAO.ExecSclr(_DAL.GetTmpTableName)
            DAO.ExecNqry(_DAL.CreateTempXLSTable(strViewXLSName))
            For Each rw As DataRow In dsSource.Tables(tbExcelName).Rows
                DAO.ExecNqry(_DAL.InsDataToTempTable(strViewXLSName), New Object() {
                             rw.Item("CustID"), rw("VODACCOUNTID"),
                             rw.Item("FaciSNO")})

            Next
        End If
        ResultViewName = "TMP_" & DAO.ExecSclr(_DAL.GetTmpTableName)
        Dim MVodIdWhere As String, ExcelWhere1 As String,
            ExcelWhere2 As String, SQLWhere As String = Nothing

        Try
            SQLWhere = _DAL.GetWhereString(dsSource, strViewXLSName)
            MVodIdWhere = _DAL.GetMvodIdWhere(dsSource.Tables(ScreenParaTableName).Rows(0).Item("MVodId").ToString)
            ExcelWhere1 = _DAL.GetExcelWhere1(strViewXLSName,
                                            dsSource.Tables(ScreenParaTableName).Rows(0).Item("ForceAmount"),
                                            dsSource.Tables(ScreenParaTableName).Rows(0).Item("Para35"))
            ExcelWhere2 = _DAL.GetExcelWhere2(strViewXLSName)
            '        CREATE VIEW "MyUser"."MyView" AS

            'Select MyTable.name

            'FROM    MyUser.MyTable

            'WHERE    MyTable.number = 'A123';
            DAO.ExecNqry("Create View " & ResultViewName & " As " & _DAL.GetSQL(dsSource, MVodIdWhere, SQLWhere, ExcelWhere1, ExcelWhere2))
            Using dtResult As DataTable = DAO.ExecQry(_DAL.getResultTable(ResultViewName))
                isVodData = dtResult.AsEnumerable.Count(Function(rw As DataRow)
                                                            Return Integer.Parse(rw("Flag").ToString) = 0
                                                        End Function)
                noVodData = dtResult.AsEnumerable.Count(Function(rw As DataRow)
                                                            Return Integer.Parse(rw("Flag").ToString) = 1
                                                        End Function)
                dtResult.TableName = "Result"
                dsResult.Tables.Add(dtResult.Copy)
                dtResult.Dispose()
            End Using
            result.ResultDataSet = dsResult.Copy

            result.ResultXML = ResultViewName
            If Not String.IsNullOrEmpty(strViewXLSName) Then
                result.ResultXML = result.ResultXML & ";" & strViewXLSName
            End If
            result.ResultXML = result.ResultXML & "," & noVodData.ToString & "," & isVodData.ToString
            'Using dtResult As DataTable = DAO.ExecQry(_DAL.GetSQL(dsSource, MVodIdWhere, SQLWhere, ExcelWhere1, ExcelWhere2))
            '    dtResult.TableName = "Result"
            '    dsResult.Tables.Add(dtResult.Copy)
            '    dtResult.Dispose()
            'End Using
            'Return dsResult
        Catch ex As Exception
            'Throw ex
            result.ErrorCode = -1
            result.ResultBoolean = False
            result.ErrorMessage = ex.ToString
        Finally
            If dsSource IsNot Nothing Then
                dsSource.Dispose()
                dsSource = Nothing
            End If
            Try
                'If Not String.IsNullOrEmpty(strViewXLSName) Then
                '    DAO.ExecNqry(_DAL.DeleteTmpTable(strViewXLSName))
                'End If

            Finally

            End Try
        End Try
        Return result
    End Function
    Public Function ImportExcel(ByVal xlsFileName As String) As DataSet
        Dim CurrentDir As String = String.Format("{0}\xls\", CableSoft.BLL.Utility.Utility.GetCurrentDirectory())
        Dim tbExcel As New DataTable(tbExcelName)
        Dim tbImport As DataTable = Nothing
        Dim dsResult As New DataSet
        tbImport = CableSoft.Utility.Heterogeneous.xls.ToDataTable(String.Format("{0}\{1}", CurrentDir, xlsFileName), , , , 1)
        tbExcel.TableName = "ImportExcel"

        Try
            tbExcel.Columns.Add(New DataColumn("CUSTID", GetType(Integer)))
            tbExcel.Columns.Add(New DataColumn("VODACCOUNTID", GetType(String)))
            tbExcel.Columns.Add(New DataColumn("FACISNO", GetType(String)))
            For Each rw As DataRow In tbImport.Rows
                If (Not rw.IsNull(Lang.ExcelCustId)) AndAlso
                        (Not rw.IsNull(Lang.ExcelVODAccountId)) AndAlso
                        (Not rw.IsNull(Lang.ExcelFaciSNo)) Then
                    Dim rwNew As DataRow = tbExcel.NewRow
                    rwNew("CUSTID") = rw.Item(Lang.ExcelCustId)
                    rwNew("VODACCOUNTID") = rw.Item(Lang.ExcelVODAccountId)
                    rwNew("FACISNO") = rw.Item(Lang.ExcelFaciSNo)
                    tbExcel.Rows.Add(rwNew)
                    tbExcel.AcceptChanges()
                End If
            Next
            dsResult.Tables.Add(tbExcel.Copy)

        Catch ex As Exception
            Throw
        Finally
            If tbImport IsNot Nothing Then
                tbImport.Dispose()
                tbImport = Nothing
            End If
            If tbExcel IsNot Nothing Then
                tbExcel.Dispose()
                tbExcel = Nothing
            End If
            Try
                System.IO.File.Delete(String.Format("{0}\{1}", CurrentDir, xlsFileName))
            Finally

            End Try
        End Try

        Return dsResult
    End Function
    Public Function GetReportParams(dsConditions As DataSet) As DataSet
        Dim CustIdListSQL As String = ""
        '取得參數
        Dim dtCondis As DataTable = dsConditions.Tables("Conditions")
        Dim dtRpt As DataTable = dsConditions.Tables("Result").Copy
        dtRpt.TableName = "Rpt"
        dsConditions.Tables.Add(dtRpt.Copy)

        Return dsConditions
    End Function
    Public Function CalculateAndCreateBill(ByVal dsSource As DataSet) As DataSet
        Dim dsCalculate As DataSet = CalculateBill(dsSource).ResultDataSet
        Try
            Return CreateBillNo(dsSource)
        Catch ex As Exception
            Throw
        Finally
            dsSource.Dispose()
            dsSource = Nothing
        End Try

    End Function
    Public Function CreateBillNo(ByVal dsSource As DataSet) As DataSet
        Dim cn As DbConnection = DAO.GetConn()
        Dim trans As DbTransaction = Nothing
        Dim CSLog As CableSoft.SO.BLL.DataLog.DataLog = Nothing
        Dim blnAutoClose As Boolean = False
        Dim dsResult As New DataSet
        Dim dtSO033 As DataTable = Nothing
        Dim dtSO033VOD As DataTable = Nothing
        Dim dtSO182 As DataTable = Nothing
        Dim dtUpdSO182 As DataTable = Nothing
        Dim dtSO062 As DataTable = Nothing
        Dim aVODAccountId As String = ""
        Dim NowDate As Date = DateTime.Now
        Dim dsDefault As DataSet = Nothing        
        If DAO.Transaction IsNot Nothing Then
            trans = DAO.Transaction
        Else
            cn.ConnectionString = Me.LoginInfo.ConnectionString
            cn.Open()
            trans = cn.BeginTransaction
            DAO.Transaction = trans
            blnAutoClose = True
        End If
        DAO.AutoCloseConn = False

        Dim cmd As DbCommand = cn.CreateCommand
        Dim BllUtility As New CableSoft.SO.BLL.Utility.Utility(Me.LoginInfo, Me.DAO)
        Dim aCitemCode As Object = Nothing
        Dim aCitemName As Object = Nothing
        Dim aCMCode As Object = Nothing
        Dim aCMName As Object = Nothing
        Dim aUCCode As Object = Nothing
        Dim aUCName As Object = Nothing
        Dim aSalePointCode As Object = Nothing
        Dim aSalePointName As Object = Nothing
        Dim intPara14 As Integer = 0
        Dim ShouldAmt As Integer = 0
        cmd.Connection = cn
        cmd.Transaction = trans
        If blnAutoClose Then
            CableSoft.BLL.Utility.Utility.SetClientInfo(Me.DAO, LoginInfo.EntryId, Lang.ClientInfoString)
        End If

        '新增回傳資料的DataTable(SO033、SO033VOD、SO182)
        '----------------------------------------------------------------------------
        dtSO033 = DAO.ExecQry(_DAL.GetSO033Sechema).Copy
        dtSO033.TableName = "SO033"
        dtSO033VOD = DAO.ExecQry(_DAL.GetSO033VODSechema).Copy
        dtSO033VOD.TableName = "SO033VOD"
        dtSO182 = DAO.ExecQry(_DAL.GetSO182Sechema).Copy
        dtSO182.TableName = "SO182"
        dsResult.Tables.Add(dtSO033.Copy)
        dsResult.Tables.Add(dtSO033VOD.Copy)
        dsResult.Tables.Add(dtSO182.Copy)
        dsResult.AcceptChanges()
        '----------------------------------------------------------------------------
        Dim note As String = Lang.InsSO033Note
        Try

            Dim lstSourceRw As List(Of DataRow) = dsSource.Tables(0).AsEnumerable.Where(Function(rw As DataRow)
                                                                                            Return rw.Item("Flag") = 0
                                                                                        End Function).OrderBy(Function(rwSource As DataRow)
                                                                                                                  Return rwSource("VODACCOUNTID")
                                                                                                              End Function).ToList
            If lstSourceRw.Count = 0 Then
                Return dsResult
            End If
            Dim RunTime As New Stopwatch()

            For Each rw As DataRow In lstSourceRw
                If aVODAccountId <> rw("VODACCOUNTID").ToString Then
                    If dsDefault Is Nothing Then
                        dsDefault = QueryDefaultValue(rw.Item("ServiceType").ToString)
                        If dsDefault.Tables(tbDefaultCitemName).Rows.Count > 0 Then
                            aCitemCode = dsDefault.Tables(tbDefaultCitemName).Rows(0).Item("CodeNo")
                            aCitemName = dsDefault.Tables(tbDefaultCitemName).Rows(0).Item("Description")
                        End If
                        If dsDefault.Tables(tbDefaultCMCodeName).Rows.Count > 0 Then
                            aCMCode = dsDefault.Tables(tbDefaultCMCodeName).Rows(0).Item("CMCode")
                            aCMName = dsDefault.Tables(tbDefaultCMCodeName).Rows(0).Item("CMName")
                        End If
                        If dsDefault.Tables(tbDefaultUCCodeName).Rows.Count > 0 Then
                            aUCCode = dsDefault.Tables(tbDefaultUCCodeName).Rows(0).Item("CodeNo")
                            aUCName = dsDefault.Tables(tbDefaultUCCodeName).Rows(0).Item("Description")
                        End If
                        intPara14 = Integer.Parse(dsDefault.Tables(tbChargeParaName).Rows(0).Item("Para14"))
                        If dsDefault.Tables(tbDefaultSalePointCodeName).Rows.Count > 0 Then
                            aSalePointCode = dsDefault.Tables(tbDefaultSalePointCodeName).Rows(0).Item("CodeNo")
                            aSalePointName = dsDefault.Tables(tbDefaultSalePointCodeName).Rows(0).Item("Description")
                        End If
                    End If
                    '新增SO033
                    Using tbSO014 As DataTable = DAO.ExecQry(_DAL.QueryDefaultSO014(intPara14),
                                                                New Object() {rw.Item("CustId"), LoginInfo.CompCode})
                        Dim aBillNo As String = BllUtility.GetInvoiceNo(Utility.InvoiceType.Temp, rw.Item("ServiceType").ToString)
                        ShouldAmt = Integer.Parse(rw("MUSTPAYCREDIT")) - Integer.Parse(rw("OverCredit")) - Integer.Parse(rw("UnPay"))                        
                        With tbSO014.Rows(0)
                            RunTime.Start()
                            DAO.ExecNqry(_DAL.InsertSO033, New Object() {
                                rw.Item("CustId"), Me.LoginInfo.CompCode, aBillNo, aCitemCode,
                                rw.Item("ShouldDate"), ShouldAmt, ShouldAmt, 0, 0, aCMCode, aCMName,
                                aUCCode, aUCName, 1, "現金", tbSO014.Rows(0).Item("ClassCode1"), rw.Item("SEQNO"), rw.Item("FaciSNo"),
                                rw.Item("ServiceType"), Lang.InsSO033Note, NowDate,
                                CableSoft.BLL.Utility.DateTimeUtility.GetDTString(NowDate),
                                Me.LoginInfo.EntryId, Me.LoginInfo.EntryName, .Item("AddrNo"),
                                .Item("StrtCode"), .Item("MduId"), .Item("ServCode"), .Item("ClctAreaCode"), .Item("ClctEn"), .Item("ClctName"),
                                .Item("AreaCode"), .Item("ClctEn"), .Item("ClctName"), aSalePointCode, aSalePointName,
                                rw("EndDate"), aCitemName, NowDate})
                            'cmd.CommandText = String.Format(_DAL.InsertSO033Test,
                            '                                rw.Item("CustId"), Me.LoginInfo.CompCode,
                            '                                aBillNo, aCitemCode,
                            '                                aCitemName,
                            '                                Date.Parse(rw.Item("ShouldDate")).ToString("yyyy/MM/dd"),
                            '                                ShouldAmt, ShouldAmt, 0, 0, aCMCode, aCMName, aUCCode, aUCName,
                            '                                1, "現金", tbSO014.Rows(0).Item("ClassCode1"), rw.Item("SEQNO"),
                            '                                rw.Item("FaciSNo"), rw.Item("ServiceType").ToString)
                            'cmd.ExecuteNonQuery()

                            RunTime.Stop()
                            Debug.Print(RunTime.Elapsed.TotalSeconds)
                            RunTime.Reset()
                        End With
                        Dim SO033VODSeqNo As String = "'-1'"
                        Dim SO182SeqNo As String = "-1"
                        Using dr As DbDataReader = DAO.ExecDtRdr(_DAL.QuerySO033VODSeqNo(rw.Item("VODAccountId"), rw.Item("EndDate")))
                            If dr.HasRows Then
                                While dr.Read
                                    If String.IsNullOrEmpty(SO033VODSeqNo) Then
                                        SO033VODSeqNo = String.Format("'{0}'", dr.Item("SEQNO"))
                                    Else
                                        SO033VODSeqNo = String.Format("{0},'{1}'", SO033VODSeqNo, dr.Item("SEQNO"))
                                    End If
                                End While
                            End If
                            dr.Close()
                        End Using
                        Using dr As DbDataReader = DAO.ExecDtRdr(_DAL.QuerySO182SeqNo(rw.Item("VODAccountId"), SO033VODSeqNo))
                            If dr.HasRows Then
                                While dr.Read
                                    If String.IsNullOrEmpty(SO182SeqNo) Then
                                        SO182SeqNo = String.Format("{0}", dr.Item("SEQNO"))
                                    Else
                                        SO182SeqNo = String.Format("{0},{1}", SO182SeqNo, dr.Item("SEQNO"))
                                    End If
                                End While
                            End If
                            dr.Close()
                        End Using

                        'dtUpdSO182 = DAO.ExecQry(_DAL.QueryUpdSO182Data(
                        '                         _DAL.QuerySO182SeqNo(rw.Item("VODAccountId"), rw.Item("EndDate"))))
                        dtUpdSO182 = DAO.ExecQry(_DAL.QueryUpdSO182Data(
                                                SO182SeqNo))
                        '更新SO182
                        'DAO.ExecNqry(_DAL.UpdateSO182(_DAL.QuerySO182SeqNo(rw.Item("VODAccountId"),
                        '                                               rw.Item("EndDate"))), New Object() {NowDate,
                        '                                                                                   Me.LoginInfo.EntryId,
                        '                                                                                   Me.LoginInfo.EntryName,
                        '                                                                                   Me.LoginInfo.EntryName,
                        '                                                                                    CableSoft.BLL.Utility.DateTimeUtility.GetDTString(NowDate),
                        '                                                                                 aBillNo})
                        DAO.ExecNqry(_DAL.UpdateSO182(SO182SeqNo), New Object() {NowDate,
                                                                                                           Me.LoginInfo.EntryId,
                                                                                                           Me.LoginInfo.EntryName,
                                                                                                           Me.LoginInfo.EntryName,
                                                                                                            CableSoft.BLL.Utility.DateTimeUtility.GetDTString(NowDate),
                                                                                                         aBillNo})
                        '更新SO033VOD
                        'DAO.ExecNqry(_DAL.UpdateSO033VOD(_DAL.QuerySO033VODSeqNo(rw.Item("VODAccountId"),
                        '                                               rw.Item("EndDate"))), New Object() {aBillNo})
                        DAO.ExecNqry(_DAL.UpdateSO033VOD(SO033VODSeqNo), New Object() {aBillNo})
                        '------------------------------------------------------------------------------------------------------------------------
                        '將新增的SO033資料新增至回傳資料
                        Using dtInsSO033 As DataTable = DAO.ExecQry(_DAL.QueryInsSO033Data, New Object() {aBillNo})
                            For Each rwIns As DataRow In dtInsSO033.Rows
                                Dim rwNew As DataRow = dsResult.Tables(dtSO033.TableName).NewRow
                                rwNew.ItemArray = rwIns.ItemArray
                                dsResult.Tables(dtSO033.TableName).Rows.Add(rwNew)
                            Next

                        End Using
                        '將Update的SO182與SO033VOD資料新增至回傳資料
                        For Each rwSO182 As DataRow In dtUpdSO182.Rows
                            dsResult.Tables(dtSO182.TableName).Rows.Add(rwSO182.ItemArray)
                            Using tbUpdSO033VOD As DataTable = DAO.ExecQry(_DAL.QueryUpdSO033VODData,
                                                                                                        New Object() {rwSO182.Item("RowId"), aBillNo})
                                For Each rwSO033VOD As DataRow In tbUpdSO033VOD.Rows
                                    Dim rwNew As DataRow = dsResult.Tables(dtSO033VOD.TableName).NewRow
                                    rwNew.ItemArray = rwSO033VOD.ItemArray
                                    dsResult.Tables(dtSO033VOD.TableName).Rows.Add(rwNew)
                                Next
                            End Using
                        Next
                        '------------------------------------------------------------------------------------------------------------------------
                    End Using
                End If
                aVODAccountId = rw.Item("VODAccountId")
            Next
            Dim paraNote As String = String.Format(Lang.paraNote, Me.LoginInfo.CompCode,
                                                            lstSourceRw.Item(0).Item("ServiceType").ToString,
                                                            Date.Parse(lstSourceRw.Item(0).Item("EndDate").ToString).ToString("yyyy/MM/dd"),
                                                            IIf(lstSourceRw.Count = 1, lstSourceRw.Item(0).Item("CustId").ToString, ""),
                                                            IIf(lstSourceRw.Count = 1, lstSourceRw.Item(0).Item("SeqNo").ToString, ""))
            '更新或新增SO062記錄
            If DAO.ExecSclr(_DAL.IsExistsSO062, New Object() {
                                                            Me.LoginInfo.CompCode,
                                                            lstSourceRw.Item(0).Item("ServiceType")}) = 0 Then
                DAO.ExecNqry(_DAL.InsSO062, New Object() {lstSourceRw.Item(0).Item("EndDate"),
                                                            Me.LoginInfo.EntryName,
                                                          CableSoft.BLL.Utility.DateTimeUtility.GetDTString(NowDate),
                                                          paraNote, Me.LoginInfo.CompCode, lstSourceRw.Item(0).Item("ServiceType")})



            Else

                DAO.ExecNqry(_DAL.UpdateSO062, New Object() {lstSourceRw.Item(0).Item("EndDate"),
                                                         Me.LoginInfo.EntryName,
                                                          CableSoft.BLL.Utility.DateTimeUtility.GetDTString(NowDate),
                                                          paraNote, Me.LoginInfo.CompCode, lstSourceRw.Item(0).Item("ServiceType")})



            End If
            dtSO062 = DAO.ExecQry(_DAL.QuerySO062Data, New Object() {lstSourceRw.Item(0).Item("ServiceType"),
                                                                   Me.LoginInfo.CompCode})
            dtSO062.TableName = "SO062"
            
            dsResult.Tables.Add(dtSO062.Copy)
            If blnAutoClose Then
                trans.Commit()
                'trans.Rollback()
            End If
            dsResult.AcceptChanges()
        Catch ex As Exception
            If blnAutoClose Then
                trans.Rollback()
            End If
            Throw ex
        Finally
            If dtSO062 IsNot Nothing Then
                dtSO062.Dispose()
                dtSO062 = Nothing
            End If
            If dtSO033 IsNot Nothing Then
                dtSO033.Dispose()
                dtSO033 = Nothing
            End If
            If dtSO033VOD IsNot Nothing Then
                dtSO033VOD.Dispose()
                dtSO033VOD = Nothing
            End If
            If dtSO182 IsNot Nothing Then
                dtSO182.Dispose()
                dtSO182 = Nothing
            End If
            If dtUpdSO182 IsNot Nothing Then
                dtUpdSO182.Dispose()
                dtUpdSO182 = Nothing
            End If
            If BllUtility IsNot Nothing Then
                BllUtility.Dispose()
                BllUtility = Nothing
            End If
            If cmd IsNot Nothing Then
                cmd.Dispose()
                cmd = Nothing
            End If
            If blnAutoClose Then
                CableSoft.BLL.Utility.Utility.ClearClientInfo(Me.DAO)
                If trans IsNot Nothing Then
                    trans.Dispose()
                    trans = Nothing
                End If
                If cn IsNot Nothing Then
                    cn.Close()
                    cn.Dispose()
                    cn = Nothing
                End If
                If blnAutoClose Then
                    DAO.AutoCloseConn = True
                End If
                If CSLog IsNot Nothing Then
                    CSLog.Dispose()
                    CSLog = Nothing
                End If
            End If

        End Try
        Return dsResult
    End Function
    Public Function QueryCompCode() As DataTable
        Return DAO.ExecQry(_DAL.QueryCompCode, New Object() {Me.LoginInfo.EntryId})
        'If Me.LoginInfo.GroupId = "0" AndAlso 1 = 0 Then
        '    Return DAO.ExecQry(_DAL.QueryCompCode("0"))
        'Else
        '    Return DAO.ExecQry(_DAL.QueryCompCode(Me.LoginInfo.GroupId), New Object() {Me.LoginInfo.EntryId})
        'End If
    End Function
    Public Function QueryMVodId(ByVal VodAccountIdS As String) As String
        Dim Result As String = Nothing
        If Not String.IsNullOrEmpty(VodAccountIdS) Then
            Using dr As DbDataReader = DAO.ExecDtRdr(_DAL.QueryMVodId(VodAccountIdS))
                While dr.Read
                    If Not String.IsNullOrEmpty(Result) Then
                        Result = String.Format("{0},{1}", Result, dr.Item("MvodId").ToString)
                    Else
                        Result = String.Format("{0}", dr.Item("MvodId").ToString)
                    End If
                End While
            End Using
        End If

        Return Result
    End Function
    Public Function QueryDefaultValue(ByVal ServiceType As String) As DataSet

        Dim dsReturn As New DataSet
        Using obj As New CableSoft.SO.BLL.Utility.Utility(Me.LoginInfo, Me.DAO)
            Dim dtPriv As DataTable = obj.GetPriv(Me.LoginInfo.EntryId, "SO32E02").Copy
            dtPriv.TableName = tbPrivName
            dsReturn.Tables.Add(dtPriv)
        End Using
        Dim dtServiceType As DataTable = DAO.ExecQry(_DAL.QueryServiceType).Copy
        dtServiceType.TableName = tbServiceTypeName
        Dim dtClosePara As DataTable = DAO.ExecQry(_DAL.QueryClosePara(),
                                                   New Object() {Me.LoginInfo.CompCode, ServiceType}).Copy
        dtClosePara.TableName = tbCloseParaName
        Dim dtChargePara As DataTable = DAO.ExecQry(_DAL.QueryChargePara(),
                                                    New Object() {Me.LoginInfo.CompCode, ServiceType}).Copy
        dtChargePara.TableName = tbChargeParaName
        Dim dtDefaultCitem As DataTable = DAO.ExecQry(_DAL.QueryDefaultCitem(), New Object() {ServiceType}).Copy
        dtDefaultCitem.TableName = tbDefaultCitemName
        Dim dtDefaultCMCode As DataTable = DAO.ExecQry(_DAL.QueryDefaultCMCode(), New Object() {ServiceType}).Copy
        dtDefaultCMCode.TableName = tbDefaultCMCodeName
        Dim dtDefaultUCCode As DataTable = DAO.ExecQry(_DAL.QueryDefaultUCCode, New Object() {ServiceType}).Copy
        dtDefaultUCCode.TableName = tbDefaultUCCodeName
        Dim dtDefaultSalePointCode As DataTable = DAO.ExecQry(_DAL.QueryDefaultSalePointCode,
                                                              New Object() {Me.LoginInfo.CompCode}).Copy

        Dim dtCompCode As DataTable = QueryCompCode.Copy
        dtCompCode.TableName = tbCompCodeName
        dtDefaultSalePointCode.TableName = tbDefaultSalePointCodeName
        dsReturn.Tables.Add(dtServiceType)
        dsReturn.Tables.Add(dtChargePara)
        dsReturn.Tables.Add(dtClosePara)
        dsReturn.Tables.Add(dtDefaultCitem)
        dsReturn.Tables.Add(dtDefaultCMCode)
        dsReturn.Tables.Add(dtDefaultSalePointCode)
        dsReturn.Tables.Add(dtDefaultUCCode)
        dsReturn.Tables.Add(dtCompCode)
        Return dsReturn
    End Function


#Region "IDisposable Support"
    Private disposedValue As Boolean ' 偵測多餘的呼叫

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                If (Me.MustDispose) AndAlso (Me.DAO IsNot Nothing) Then
                    DAO.Dispose()
                End If
                If Lang IsNot Nothing Then
                    Lang.Dispose()
                    Lang = Nothing
                End If
                ' TODO: 處置 Managed 狀態 (Managed 物件)。
            End If

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
