Imports System.Data.Common
Imports CableSoft.BLL.BillingAPI
Imports CableSoft.BLL.Utility
Public Class BatchCreate
    Inherits BLLBasic
    Implements IDisposable

    Private _DAL As New BatchCreateDALMultiDB(Me.LoginInfo.Provider)
    Private Language As New CableSoft.BLL.Language.SO61.Invoice
    Private executeTime As Date = Date.Now
    Public Sub New()
    End Sub
    Public Sub New(ByVal LoginInfo As CableSoft.BLL.Utility.LoginInfo)
        MyBase.New(LoginInfo)
    End Sub
    Public Sub New(ByVal LoginInfo As CableSoft.BLL.Utility.LoginInfo, ByVal DAO As CableSoft.Utility.DataAccess.DAO)
        MyBase.New(LoginInfo, DAO)
    End Sub
    Public Sub New(ByVal LoginInfo As CableSoft.BLL.Utility.LoginInfo, ByVal DBConnection As System.Data.Common.DbConnection)
        MyBase.New(LoginInfo, DBConnection)
    End Sub

    Public Function QueryExceptInvDetail(ByVal SEQ As String, ByVal DataType As Integer) As DataSet
        Dim ds As New DataSet
        Try
            Using tbDetail As DataTable = DAO.ExecQry(_DAL.QueryExceptInvDetail(DataType), New Object() {SEQ})
                tbDetail.TableName = "DETAIL"
                ds.Tables.Add(tbDetail.Copy)
            End Using
            Return ds.Copy
        Catch ex As Exception
            Throw ex
        Finally
            If ds IsNot Nothing Then
                ds.Dispose()
                ds = Nothing
            End If
        End Try

    End Function
    Public Function Execute(ByVal scrDataSet As DataSet) As RIAResult
        Dim cn As DbConnection = DAO.GetConn()
        Dim trans As DbTransaction = Nothing
        Dim blnAutoClose As Boolean = False
        Dim tbINV001 As DataTable = Nothing
        Dim tbINV063 As DataTable = Nothing
        Dim aPrefixString As String = Nothing
        Dim sw As New Stopwatch
        Dim result As New RIAResult With {.ErrorCode = -1, .ErrorMessage = "init", .ResultBoolean = False}
        Dim ds As New DataSet
        Dim sfParameter As String = Nothing
        If DAO.Transaction IsNot Nothing Then
            trans = DAO.Transaction
        Else
            If cn.State <> ConnectionState.Open Then
                cn.ConnectionString = Me.LoginInfo.ConnectionString
                cn.Open()
            End If
            trans = cn.BeginTransaction
            DAO.Transaction = trans
            blnAutoClose = True
        End If
        DAO.AutoCloseConn = False
        sw.Start()
        tbINV001 = DAO.ExecQry(_DAL.QueryINV001, New Object() {LoginInfo.CompCode.ToString})
        tbINV063 = DAO.ExecQry(_DAL.QueryINV063, New Object() {LoginInfo.CompCode})
        For Each rw As DataRow In scrDataSet.Tables("INVSELECTED").Rows
            If String.IsNullOrEmpty(aPrefixString) Then
                aPrefixString = rw.Item("Prefix") & Right("00000000" & rw("StartNum"), 8)
            Else
                aPrefixString = aPrefixString & "," & rw.Item("Prefix") & Right("00000000" & rw("StartNum"), 8)
            End If
        Next
        Try
            Dim InPara As New Dictionary(Of String, Object)
            Dim outPara As New Dictionary(Of String, Object)
            Dim retVal As Object = Nothing
            outPara.Add("p_RetCode", Nothing)
            outPara.Add("p_RetMsg", Nothing)
            outPara.Add("p_LogDateTime", Nothing)
            InPara.Add("p_User", LoginInfo.EntryName)
            InPara.Add("p_CompId", LoginInfo.CompCode)
            InPara.Add("p_LinkToMis", "Y")
            InPara.Add("p_DbLink", Nothing)
            InPara.Add("p_HowToCreate", Integer.Parse(scrDataSet.Tables("CONDITION").Rows(0).Item("HOWTOCREATE")))
            InPara.Add("p_InvDateEqualToChargeDate", scrDataSet.Tables("CONDITION").Rows(0).Item("SAMEINVDATE"))
            If (DBNull.Value.Equals(scrDataSet.Tables("CONDITION").Rows(0).Item("INVDATE"))) Then
                InPara.Add("p_InvDate", Nothing)
            Else
                InPara.Add("p_InvDate", Date.Parse(scrDataSet.Tables("CONDITION").Rows(0).Item("INVDATE")).ToString("yyyy/MM/dd"))
            End If

            InPara.Add("p_InvYearMonth", scrDataSet.Tables("CONDITION").Rows(0).Item("YEARMONTH"))
            InPara.Add("p_ChargeStartdate", Date.Parse(scrDataSet.Tables("CONDITION").Rows(0).Item("INVSTARTDATE")).ToString("yyyy/MM/dd"))
            InPara.Add("p_ChargeStopDate", Date.Parse(scrDataSet.Tables("CONDITION").Rows(0).Item("INVENDDATE")).ToString("yyyy/MM/dd"))
            InPara.Add("p_IdentifyID1", "1")
            InPara.Add("p_IdentifyID2", 0)
            InPara.Add("p_SystemID", tbINV001.Rows(0).Item("CheckInvNum"))
            InPara.Add("p_PrefixString", aPrefixString)
            InPara.Add("p_OrderBy", Integer.Parse(scrDataSet.Tables("CONDITION").Rows(0).Item("ORDERNUM")))
            InPara.Add("p_MisDbOwner", tbINV001.Rows(0).Item("MisOwner").ToString & "")
            InPara.Add("p_ShowFaci", Integer.Parse(tbINV001.Rows(0).Item("ShowFaci")))
            InPara.Add("p_StarCMTVMail", Integer.Parse("0" & tbINV001.Rows(0).Item("StartTVMail")))
            InPara.Add("p_FilterBusinessId", 3)
            InPara.Add("p_FilterInvoiceKind", Integer.Parse(scrDataSet.Tables("CONDITION").Rows(0).Item("INVOICEKIND")))
            'Using tb As DataTable = DAO.ExecQry(_DAL.exesf, New Object() {InPara.Item("p_User"), InPara.Item("p_CompId"), InPara.Item("p_LinkToMis"), InPara.Item("p_DbLink"), InPara.Item("p_HowToCreate"), InPara.Item("p_InvDateEqualToChargeDate") _
            '                                    , InPara.Item("p_InvDate"), InPara.Item("p_InvYearMonth"), InPara.Item("p_ChargeStartdate"), InPara.Item("p_ChargeStopDate"), InPara.Item("p_IdentifyID1"), InPara.Item("p_IdentifyID2"), InPara.Item("p_SystemID"), InPara.Item("p_PrefixString"), InPara.Item("p_OrderBy") _
            '                                    , InPara.Item("p_MisDbOwner"), InPara.Item("p_ShowFaci"), InPara.Item("p_StarCMTVMail"), InPara.Item("p_FilterBusinessId"), InPara.Item("p_FilterInvoiceKind")})
            '    If tb.Rows.Count = 1 Then

            '    End If
            'End Using
            sfParameter = ""
            For Each sKey As String In InPara.Keys
                If InPara.Item(sKey) Is Nothing Then
                    sfParameter = sfParameter & ",null"
                Else

                    If InPara.Item(sKey).GetType Is GetType(Integer) Then
                        sfParameter = sfParameter & "," & InPara(sKey)
                    Else
                        sfParameter = sfParameter & ",'" & InPara(sKey) & "'"
                    End If
                End If
            Next
            result.Message2 = sfParameter.Substring(1)
            If Not DAO.ExecSF(DAO.GetConn, "SF_ASSIGNINVID", InPara, outPara, retVal) Then
                result.ResultBoolean = False
                result.ErrorCode = -99
                result.Message = "Excute store function failed"
                Return result
            End If

            If (Not String.IsNullOrEmpty(outPara("p_LogDateTime"))) Then
                Using tbUnusual As DataTable = DAO.ExecQry(_DAL.QueryUnusualInv, New Object() {
                                                                LoginInfo.CompCode.ToString,
                                                               Date.Parse(outPara("p_LogDateTime"))
                                                           })
                    tbUnusual.TableName = "UNUSUAL"
                    ds.Tables.Add(tbUnusual.Copy)
                End Using
                'Using tbUnusual As DataTable = DAO.ExecQry(_DAL.QueryUnusualInv)
                '    tbUnusual.TableName = "UNUSUAL"
                '    ds.Tables.Add(tbUnusual.Copy)
                'End Using

            End If
            If (Integer.Parse(outPara("p_RetCode") & "") <> 0) Then
                result.ErrorCode = Integer.Parse(outPara("p_RetCode") & "")
                result.ErrorMessage = "Execute SF_ASSIGNINVID Error"
                result.ResultBoolean = False
                If (Not String.IsNullOrEmpty(outPara("p_RetMsg"))) Then
                    result.ErrorMessage = outPara("p_RetMsg")
                End If
                trans.Rollback()
                Return result
            Else
                Using tbAllData As DataTable = DAO.ExecQry(_DAL.QueryAllCreateData, New Object() {LoginInfo.CompCode.ToString})
                    tbAllData.TableName = "ALLDATA"
                    ds.Tables.Add(tbAllData.Copy)
                End Using
                Using tbUseInv As DataTable = DAO.ExecQry(_DAL.QueryStarttoEndInv, New Object() {LoginInfo.CompCode.ToString})
                    tbUseInv.TableName = "USEINV"
                    ds.Tables.Add(tbUseInv.Copy)
                End Using
                If (Integer.Parse("0" & tbINV001.Rows(0).Item("StartEinvoice")) = 1) Then
                    Using tbElectricInv As DataTable = DAO.ExecQry(_DAL.QueryElectricInvData, New Object() {LoginInfo.CompCode.ToString})
                        tbElectricInv.TableName = "ELECTRIC"
                        ds.Tables.Add(tbElectricInv.Copy)
                    End Using
                    Using tbNoneElectricInv As DataTable = DAO.ExecQry(_DAL.QueryNoneElectricInvData, New Object() {LoginInfo.CompCode.ToString})
                        tbNoneElectricInv.TableName = "NONEELECTRIC"
                        ds.Tables.Add(tbNoneElectricInv.Copy)
                    End Using
                End If
            End If
            result.ErrorCode = 0
            result.ErrorMessage = Nothing
            result.ResultBoolean = True
            result.Message = String.Format(Language.CreateResult, Math.Round(sw.Elapsed.TotalSeconds, 2))

            result.ResultDataSet = ds.Copy
            trans.Commit()
            Return result
        Catch ex As Exception
            If blnAutoClose Then
                trans.Rollback()
            End If
            Throw ex
        Finally
            If ds IsNot Nothing Then
                ds.Dispose()
                ds = Nothing
            End If
            If tbINV001 IsNot Nothing Then
                tbINV001.Dispose()
                tbINV001 = Nothing
            End If
            If tbINV063 IsNot Nothing Then
                tbINV063.Dispose()
                tbINV063 = Nothing
            End If
            If sw IsNot Nothing Then
                sw.Stop()
                sw = Nothing
            End If
        End Try

    End Function
    Public Function GetCompCode() As DataTable
        Try


            Return DAO.ExecQry(_DAL.GetCompCode("1",
                                                 CableSoft.BLL.Utility.Utility.GetCompanyTableName(Me.LoginInfo, Me.DAO),
                                                    CableSoft.BLL.Utility.Utility.GetLoginTableName),
                                New Object() {Me.LoginInfo.EntryId})
        Catch ex As Exception
            Throw
        End Try



    End Function
    Public Function QueryExceptInvInfo(ByVal scrDataset As DataSet, ByVal DataType As Integer) As DataSet
        Dim ds As New DataSet
        Dim invoiceKind As Integer = 2
        Dim SEQ As String = "X"
        Try
            If Not DBNull.Value.Equals(scrDataset.Tables("CONDITION").Rows(0).Item("InvoiceKind")) Then
                invoiceKind = Integer.Parse(scrDataset.Tables("CONDITION").Rows(0).Item("InvoiceKind"))
            End If
            Using tbInfo As DataTable = DAO.ExecQry(_DAL.QueryExceptInvInfo(scrDataset.Tables(0).Rows(0).Item("OrderNum"),
                                                                            invoiceKind, DataType), New Object() {
                                                            LoginInfo.CompCode.ToString,
                                                             scrDataset.Tables(0).Rows(0).Item("INVSTARTDATE"),
                                                      scrDataset.Tables(0).Rows(0).Item("InvEndDate"),
                                                      scrDataset.Tables(0).Rows(0).Item("HowToCreate"),
                                                      invoiceKind
                                                    }).Copy
                tbInfo.TableName = "MASTER"
                If tbInfo.Rows.Count > 0 Then
                    SEQ = tbInfo.Rows(0).Item("SEQ")
                End If
                Using dsDetail As DataSet = QueryExceptInvDetail(SEQ, DataType)
                    ds.Tables.Add(dsDetail.Tables(0).Copy)
                End Using
                ds.Tables.Add(tbInfo)

            End Using

            Return ds.Copy
        Catch ex As Exception
            Throw ex
        Finally
            If ds IsNot Nothing Then
                ds.Dispose()
                ds = Nothing
            End If

        End Try

    End Function
    Public Function QueryCanCreateInv(ByVal scrDataset As DataSet) As DataSet
        Dim ds As New DataSet
        Dim iExceptCount As Integer = 0
        Dim iAvaliableCount As Integer = 0
        Dim dtResult As DataTable = New DataTable("RESULT")
        Dim invoiceKind As Integer = 2
        dtResult.Columns.Add(New DataColumn("EXCEPTCOUNT", GetType(Integer)))
        dtResult.Columns.Add(New DataColumn("AVAILABLECOUNT", GetType(Integer)))
        dtResult.Columns.Add(New DataColumn("NOCOUNT", GetType(Integer)))
        dtResult.Columns.Add(New DataColumn("SALEAMOUNT1", GetType(Integer)))
        dtResult.Columns.Add(New DataColumn("TAXAMOUNT1", GetType(Integer)))
        dtResult.Columns.Add(New DataColumn("INVAMOUNT1", GetType(Integer)))
        dtResult.Columns.Add(New DataColumn("SALEAMOUNT2", GetType(Integer)))
        dtResult.Columns.Add(New DataColumn("TAXAMOUNT2", GetType(Integer)))
        dtResult.Columns.Add(New DataColumn("INVAMOUNT2", GetType(Integer)))
        dtResult.Columns.Add(New DataColumn("ERRMSG", GetType(String)))

        Dim rwResult As DataRow = dtResult.NewRow
        Try
            If Not DBNull.Value.Equals(scrDataset.Tables("CONDITION").Rows(0).Item("InvoiceKind")) Then
                invoiceKind = Integer.Parse(scrDataset.Tables("CONDITION").Rows(0).Item("InvoiceKind"))
            End If
            Dim aAutoCreateNum As Integer = DAO.ExecSclr(_DAL.QueryAutoCreateNum, New Object() {LoginInfo.CompCode.ToString})
            Using tbExceptCount As DataTable = DAO.ExecQry(_DAL.QueryCanCreateInv(scrDataset.Tables("CONDITION").Rows(0).Item("InvoiceKind")),
                                                      New Object() {aAutoCreateNum, LoginInfo.CompCode.ToString,
                                                      scrDataset.Tables("CONDITION").Rows(0).Item("INVSTARTDATE"),
                                                      scrDataset.Tables("CONDITION").Rows(0).Item("InvEndDate"),
                                                      scrDataset.Tables("CONDITION").Rows(0).Item("HowToCreate"),
                                                      scrDataset.Tables("CONDITION").Rows(0).Item("InvoiceKind")}).Copy
                tbExceptCount.TableName = "ExceptCount"
                If (Not DBNull.Value.Equals(tbExceptCount.Rows(0).Item("Count"))) Then
                    iExceptCount = Integer.Parse("0" & tbExceptCount.Rows(0).Item("Count").ToString)
                End If
                rwResult("EXCEPTCOUNT") = iExceptCount
                For Each rwInv099 As DataRow In scrDataset.Tables("INVSELECTED").Rows
                    iAvaliableCount += Integer.Parse(rwInv099.Item("ENDNUM")) - Integer.Parse(rwInv099.Item("CURNUM")) + 1
                Next
                rwResult("AVAILABLECOUNT") = iAvaliableCount
                Using tbSum As DataTable = DAO.ExecQry(_DAL.QueryMustCreateAmount(invoiceKind), New Object() {
                                                        LoginInfo.CompCode.ToString,
                                                         scrDataset.Tables("CONDITION").Rows(0).Item("INVSTARTDATE"),
                                                        scrDataset.Tables("CONDITION").Rows(0).Item("InvEndDate"),
                                                        scrDataset.Tables("CONDITION").Rows(0).Item("HowToCreate"),
                                                        invoiceKind
                                                       })
                    For Each rw As DataRow In tbSum.Rows
                        rwResult("SALEAMOUNT1") = rw("SALEAMOUNT")
                        rwResult("TAXAMOUNT1") = rw("TAXAMOUNT")
                        rwResult("INVAMOUNT1") = rw("INVAMOUNT")
                    Next

                End Using
                Using tbSum2 As DataTable = DAO.ExecQry(_DAL.QueryNoCreateAmount(invoiceKind), New Object() {
                                                        LoginInfo.CompCode.ToString,
                                                         scrDataset.Tables("CONDITION").Rows(0).Item("INVSTARTDATE"),
                                                        scrDataset.Tables("CONDITION").Rows(0).Item("InvEndDate"),
                                                        scrDataset.Tables("CONDITION").Rows(0).Item("HowToCreate"),
                                                        invoiceKind
                                                       })
                    For Each rw As DataRow In tbSum2.Rows
                        rwResult("SALEAMOUNT2") = rw("SALEAMOUNT")
                        rwResult("TAXAMOUNT2") = rw("TAXAMOUNT")
                        rwResult("INVAMOUNT2") = rw("INVAMOUNT")
                        rwResult("NOCOUNT") = rw("NOCOUNT")
                    Next
                End Using
                If iExceptCount > iAvaliableCount Then
                    rwResult("ERRMSG") = String.Format(Language.CannotBatchCreate, iExceptCount, iAvaliableCount)
                End If

                dtResult.Rows.Add(rwResult.ItemArray)
                ds.Tables.Add(dtResult.Copy)
            End Using

            Return ds.Copy
        Catch ex As Exception
            Throw ex
        Finally
            If ds IsNot Nothing Then
                ds.Dispose()
                ds = Nothing
            End If
        End Try
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
                'result.ErrorMessage = Language.NoPermission
            End If
            'If Integer.Parse(DAO.ExecSclr(_DAL.chkAuthority(Me.LoginInfo.GroupId), New Object() {Mid})) = 0 Then
            '    result.ResultBoolean = False
            '    result.ErrorCode = -1
            '    result.ErrorMessage = Language.NoPermission
            '    Return result
            'End If

        Catch ex As Exception
            result.ErrorMessage = ex.ToString
            result.ResultBoolean = False
            result.ErrorCode = -2
        Finally

        End Try
        Return result

    End Function
    Public Function QueryAllData() As DataSet
        Dim ds As New DataSet
        Try
            Using tbINV001 As DataTable = DAO.ExecQry(_DAL.QueryINV001, New Object() {LoginInfo.CompCode.ToString})
                tbINV001.TableName = "INV001"
                ds.Tables.Add(tbINV001.Copy)
            End Using
            Using tbINV099 As DataTable = DAO.ExecQry(_DAL.QueryINV099, New Object() {LoginInfo.CompCode.ToString})
                tbINV099.TableName = "INV099"
                ds.Tables.Add(tbINV099.Copy)
            End Using
            Using tbInvoiceKinde As DataTable = DAO.ExecQry(_DAL.QueryInvoiceKind)
                tbInvoiceKinde.TableName = "INVOICEKIND"
                ds.Tables.Add(tbInvoiceKinde.Copy)

            End Using
            Using tbQueryGridOrder As DataTable = DAO.ExecQry(_DAL.QueryGridOrder)
                tbQueryGridOrder.TableName = "GRIDORDER"
                ds.Tables.Add(tbQueryGridOrder.Copy)
            End Using
            Using tbHowToCreate As DataTable = DAO.ExecQry(_DAL.QueryHowtoCreate)
                tbHowToCreate.TableName = "HOWTOCREATE"
                ds.Tables.Add(tbHowToCreate.Copy)
            End Using
            Using tbCompCode As DataTable = GetCompCode()
                tbCompCode.TableName = "COMPCODE"
                ds.Tables.Add(tbCompCode.Copy)
            End Using
            Return ds.Copy
        Catch ex As Exception
            Throw ex
        Finally
            If ds IsNot Nothing Then
                ds.Dispose()
                ds = Nothing
            End If
        End Try


    End Function
#Region "IDisposable Support"
    Private disposedValue As Boolean ' 偵測多餘的呼叫

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
                If DAO.AutoCloseConn Then
                    DAO.CloseConn()
                    DAO.Dispose()
                    DAO = Nothing
                End If
                ' TODO: 處置 Managed 狀態 (Managed 物件)。
            End If

            ' TODO: 釋放 Unmanaged 資源 (Unmanaged 物件) 並覆寫下方的 Finalize()。
            ' TODO: 將大型欄位設為 null。
        End If
        disposedValue = True
    End Sub

    ' TODO: 只有當上方的 Dispose(disposing As Boolean) 具有要釋放 Unmanaged 資源的程式碼時，才覆寫 Finalize()。
    'Protected Overrides Sub Finalize()
    '    ' 請勿變更這個程式碼。請將清除程式碼放在上方的 Dispose(disposing As Boolean) 中。
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' Visual Basic 加入這個程式碼的目的，在於能正確地實作可處置的模式。
    Public Sub Dispose() Implements IDisposable.Dispose
        ' 請勿變更這個程式碼。請將清除程式碼放在上方的 Dispose(disposing As Boolean) 中。
        Dispose(True)
        ' TODO: 覆寫上列 Finalize() 時，取消下行的註解狀態。
        ' GC.SuppressFinalize(Me)
    End Sub
#End Region
End Class
