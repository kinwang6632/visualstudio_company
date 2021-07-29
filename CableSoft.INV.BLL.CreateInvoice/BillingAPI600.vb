Imports System.Data.Common
Imports CableSoft.BLL.BillingAPI
Imports CableSoft.BLL.Utility
Public Class BillingAPI600
    Inherits BLLBasic
    Implements IDisposable, CableSoft.BLL.BillingAPI.IBillingAPI

    Private _DAL As New BillingAPI600DALMultiDB(Me.LoginInfo.Provider)
    Private Language As New CableSoft.BLL.Language.SO61.BillingAPI600Language
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
    Function Execute(SeqNo As Integer, InData As DataSet) As CableSoft.BLL.Utility.RIAResult Implements CableSoft.BLL.BillingAPI.IBillingAPI.Execute
        Dim result As New RIAResult
      
        Dim CompId As String = InData.Tables("Main").Rows(0).Item("Compcode")
        Dim aSysId As String = Nothing

        Dim aBusinessId As String = Nothing
        If Not DBNull.Value.Equals(InData.Tables("Inv").Rows(0).Item("BusinessId")) AndAlso
           Not String.IsNullOrEmpty(InData.Tables("Inv").Rows(0).Item("BusinessId")) Then
            aBusinessId = InData.Tables("Inv").Rows(0).Item("BusinessId")
            Dim retIdErr As String = Nothing
            If Not CableSoft.BLL.Utility.Utility.InvNoVerify(InData.Tables("Inv").Rows(0).Item("BusinessId"), retIdErr) Then
                result.ResultBoolean = False
                result.ErrorCode = -611
                result.ErrorMessage = retIdErr
                Return result
            End If
        End If
        Dim xmlText As String = DAO.ExecSclr(_DAL.QueryINV003Param, New Object() {CompId})
        If String.IsNullOrEmpty(xmlText) Then
            result.ResultBoolean = False
            result.ErrorCode = -999
            result.ErrorMessage = Language.noINV003
            Return result
        End If
        Dim oXml As New Xml.XmlDocument()
        oXml.LoadXml(xmlText)
        aSysId = oXml.DocumentElement.GetAttribute("SysID")
        oXml = Nothing
        If String.IsNullOrEmpty(aSysId) Then
            result.ResultBoolean = False
            result.ErrorCode = -999
            result.ErrorMessage = Language.noSysID
            Return result
        End If
        Dim oo As String = getInvoiceYearMonth(InData.Tables("Inv").Rows(0).Item("InvDate"))
        Dim tb099 As DataTable = DAO.ExecQry(_DAL.QueryInv099, New Object() {CompId, oo})
        If tb099.Rows.Count = 0 Then
            result.ResultBoolean = False
            result.ErrorCode = -600
            result.ErrorMessage = Language.NoInvChr
            Return result
        End If
        If tb099.Rows.Count > 1 Then
            result.ResultBoolean = False
            result.ErrorCode = -601
            result.ErrorMessage = Language.getMultiInvChr
            Return result
        End If

        Dim cn As DbConnection = DAO.GetConn()
        Dim trans As DbTransaction = Nothing
        Dim blnAutoClose As Boolean = False
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
        '#8706
        Me.LoginInfo.EntryName = InData.Tables("Inv").Rows(0).Item("Upden")

        Try
            Dim aSeq As String = DateTime.Now.ToString("yyyyMMdd") & DAO.ExecSclr(_DAL.getInvSeq()).ToString
            Dim aCustId As Object = InData.Tables("Inv").Rows(0).Item("CustID")
            Dim aInvTitle As Object = InData.Tables("Inv").Rows(0).Item("InvTitle")
            Dim aZipCode As Object = InData.Tables("Inv").Rows(0).Item("ZipCode")
            Dim aInvAddr As Object = InData.Tables("Inv").Rows(0).Item("InvAddr")
            Dim aMailAddr As Object = InData.Tables("Inv").Rows(0).Item("MailAddr")
            Dim aChargeDate As Object = InData.Tables("Inv").Rows(0).Item("ChargeDate")
            Dim aInvAmount As Object = InData.Tables("Inv").Rows(0).Item("InvAmount")
            Dim aSaleAmount As Object = InData.Tables("Inv").Rows(0).Item("InvAmount")
            Dim aTaxAmount As Object = 0
            '#8706
            'Dim aCaller As String = InData.Tables("Main").Rows(0).Item("Caller").ToString
            Dim aCaller As String = Me.LoginInfo.EntryName

            Dim aTaxRate As Integer = 0
            Dim aCustSName As Object = InData.Tables("Inv").Rows(0).Item("CustSName").ToString
            Dim aMemo1 As Object = InData.Tables("Inv").Rows(0).Item("Memo1")
            Dim aMemo2 As Object = InData.Tables("Inv").Rows(0).Item("Memo2")
            Dim aLoveNum As Object = InData.Tables("Inv").Rows(0).Item("LoveNum")
            Dim aA_CarrierId1 As Object = InData.Tables("Inv").Rows(0).Item("A_CarrierId1")
            Dim aA_CarrierId2 As Object = InData.Tables("Inv").Rows(0).Item("A_CarrierId2")
            Dim aCarrierType As Object = InData.Tables("Inv").Rows(0).Item("CarrierType")
            Dim aCarrierId1 As Object = InData.Tables("Inv").Rows(0).Item("CarrierId1")
            Dim aCarrierId2 As Object = InData.Tables("Inv").Rows(0).Item("CarrierId2")
            Dim aCusComp As Object = InData.Tables("Main").Rows(0).Item("CusComp")
            Dim aCusOwner As Object = InData.Tables("Main").Rows(0).Item("CusOwner")
            Dim aLinkToMis As String = "N"
            Dim aTaxType As String = InData.Tables("Inv").Rows(0).Item("TaxType").ToString()
            Select Case Integer.Parse(InData.Tables("Inv").Rows(0).Item("TaxType").ToString())

                Case 1
                    aTaxAmount = Math.Round(Integer.Parse(InData.Tables("Inv").Rows(0).Item("InvAmount")) * 0.05,
                                              0, MidpointRounding.AwayFromZero)
                    aTaxRate = 5
            End Select
            aSaleAmount = Integer.Parse(aInvAmount) - Integer.Parse(aTaxAmount)
            If Not DBNull.Value.Equals(aCusComp) AndAlso Not DBNull.Value.Equals(aCusOwner) Then
                aLinkToMis = "Y"
            End If


            DAO.ExecNqry(_DAL.InsertINV049, New Object() {aSeq, CompId, aCustId, aBusinessId, aInvTitle,
                                                          aZipCode, aInvAddr, aMailAddr, aSaleAmount,
                                                          aTaxAmount, aInvAmount, aCaller, aCustSName, Date.Parse(aChargeDate),
                                                          aMemo1, aMemo2, aLoveNum, aA_CarrierId1, aA_CarrierId2,
                                                          aCarrierType, aCarrierId1, aCarrierId2, 0,
                                                          InData.Tables("Inv").Rows(0).Item("TaxType").ToString(),
                                                          aSaleAmount, aTaxAmount, aInvAmount})


            Dim aDetailBillitemNo As Integer = 0
            Dim aDetailAmount As Integer = 0

            For Each row As DataRow In InData.Tables("Detail").Rows
                If aLinkToMis = "Y".ToUpper() Then
                    If DBNull.Value.Equals(row.Item("ServiceType")) OrElse String.IsNullOrEmpty(row.Item("ServiceType").ToString) Then
                        result.ResultBoolean = False
                        result.ErrorCode = -996
                        result.ErrorMessage = String.Format(Language.needServiceType, row.Item("Description"))
                        Return result
                    End If
                    If DBNull.Value.Equals(row.Item("BillIDItemNo")) OrElse String.IsNullOrEmpty(row.Item("BillIDItemNo").ToString) Then
                        result.ResultBoolean = False
                        result.ErrorCode = -995
                        result.ErrorMessage = String.Format(Language.needItem, row.Item("Description"))
                        Return result
                    End If
                End If
                Dim o As Object = DAO.ExecSclr(_DAL.QueryINV005Name, New Object() {row.Item("ItemID"), CompId})
                If o IsNot Nothing AndAlso Not DBNull.Value.Equals(o) Then
                    If o <> row.Item("Description") Then
                        result.ResultBoolean = False
                        result.ErrorCode = -998
                        result.ErrorMessage = String.Format(Language.differItemSource, row.Item("Description"))
                        Return result
                    End If
                    '判斷主稅別與明細是否相符,由前端判斷掉,所以此段不需要再檢查一次
                    'If aTaxType <> DAO.ExecSclr(_DAL.QueryINV005TaxCode, New Object() {row.Item("ItemID"), CompId}).ToString() Then
                    '    result.ResultBoolean = False
                    '    result.ErrorCode = -994
                    '    result.ErrorMessage = String.Format(Language.differItemTaxCode, row.Item("Description"))
                    '    Return result
                    'End If
                End If


                Dim aDetailQuantity As Integer = 1
                Dim aDetailTaxAmount As Integer = 0
                Dim aDetailTotalAmount As Integer = Integer.Parse(row.Item("TotalAmount").ToString)
                Dim aDetailSaleAmount = 0
                Dim aDetailStartDate As Object = Nothing
                Dim aDetailEndDate As Object = Nothing
                Dim aDetailChargeEn As String = Nothing
                aDetailAmount = aDetailAmount + aDetailTotalAmount
                If row.Item("StartDate") IsNot Nothing AndAlso Not DBNull.Value.Equals(row.Item("StartDate")) Then
                    aDetailStartDate = Date.Parse(row.Item("StartDate"))
                End If
                If row.Item("EndDate") IsNot Nothing AndAlso Not DBNull.Value.Equals(row.Item("EndDate")) Then
                    aDetailEndDate = Date.Parse(row.Item("EndDate"))
                End If
                If row.Item("ChargeEn") IsNot Nothing AndAlso Not DBNull.Value.Equals(row.Item("ChargeEn")) Then
                    aDetailChargeEn = row.Item("ChargeEn")
                End If
                If row.Item("QUANTITY") IsNot Nothing AndAlso Not String.IsNullOrEmpty(row.Item("QUANTITY").ToString) Then
                    aDetailQuantity = Integer.Parse(row.Item("QUANTITY").ToString)
                    If aDetailQuantity <= 0 Then aDetailQuantity = 1
                End If
                Select Case aTaxRate
                    Case 5
                        'aDetailTaxAmount = Math.Round(aDetailTotalAmount * 0.05,
                        '                      0, MidpointRounding.AwayFromZero)
                        aDetailSaleAmount = Math.Round(aDetailTotalAmount / 1.05,
                                              0, MidpointRounding.AwayFromZero)
                    Case 0
                        aDetailSaleAmount = aDetailTotalAmount
                End Select
                aDetailTaxAmount = aDetailTotalAmount - aDetailSaleAmount
                Dim aDetailUnitPrice As Integer = Math.Round(aDetailSaleAmount / aDetailQuantity, 2, MidpointRounding.AwayFromZero)
                Dim aRealBillitemNo As Integer = 0
                If DBNull.Value.Equals(row("BillIDItemNo")) Then
                    aDetailBillitemNo -= 1
                    aRealBillitemNo = aDetailBillitemNo
                Else
                    aRealBillitemNo = Integer.Parse(row("BillIDItemNo"))
                End If

                DAO.ExecNqry(_DAL.InsertINV050, New Object() {
                             aSeq, row("BillID"), aRealBillitemNo, Integer.Parse(InData.Tables("Inv").Rows(0).Item("TaxType").ToString()), Date.Parse(aChargeDate), _
                             row.Item("ItemID"), row("Description"), aDetailQuantity, aDetailUnitPrice, aTaxRate, _
                             aDetailTaxAmount, aDetailTotalAmount, aDetailStartDate, aDetailEndDate, aDetailChargeEn, _
                           aLinkToMis, row.Item("ServiceType").ToString
                         })

            Next

            If aInvAmount <> aDetailAmount Then
                result.ResultBoolean = False
                result.ErrorCode = -999
                result.ErrorMessage = Language.differAmount
                Return result
            End If
            Dim InPara As New Dictionary(Of String, Object)
            Dim outPara As New Dictionary(Of String, Object)
            Dim retVal As Object = Nothing
            outPara.Add("p_RetCode", Nothing)
            outPara.Add("p_RetMsg", Nothing)
            outPara.Add("p_LogDateTime", Nothing)
            InPara.Add("p_User", aCaller)
            InPara.Add("p_CompId", CompId)
            InPara.Add("p_LinkToMis", aLinkToMis)
            InPara.Add("p_DbLink", Nothing)
            InPara.Add("p_HowToCreate", 3)
            InPara.Add("p_InvDateEqualToChargeDate", 2)
            InPara.Add("p_InvDate", InData.Tables("Inv").Rows(0).Item("InvDate").ToString())
            InPara.Add("p_InvYearMonth", tb099.Rows(0).Item("YearMonth").ToString())
            InPara.Add("p_ChargeStartdate", aChargeDate)
            InPara.Add("p_ChargeStopDate", aChargeDate)
            InPara.Add("p_IdentifyID1", 1)
            InPara.Add("p_IdentifyID2", "0")
            InPara.Add("p_SystemID", aSysId)
            InPara.Add("p_PrefixString", tb099.Rows(0).Item("Prefix").ToString & tb099.Rows(0).Item("StartNum").ToString())
            InPara.Add("p_OrderBy", 1)
            InPara.Add("p_MisDbOwner", aCusOwner)
            InPara.Add("p_ShowFaci", 0)
            InPara.Add("p_StarCMTVMail", 0)
            InPara.Add("p_FilterBusinessId", 1)
            InPara.Add("p_FilterInvoiceKind", 1)
            InPara.Add("p_QrySeq", aSeq)
            InPara.Add("CusComp", aCusComp)

            If Not DAO.ExecSF(DAO.GetConn, "sf_assigninvidsingle", InPara, outPara, retVal) Then
                result.ResultBoolean = False
                result.ErrorCode = -99
                result.Message = "Excute store function failed"
                Return result
            End If
            If retVal = 0 Then
                If String.IsNullOrEmpty(outPara.Item("p_RetMsg")) Then
                    result.ErrorCode = -995
                    result.ErrorMessage = Language.noAnyInv
                    result.ResultBoolean = False
                Else
                    result.ErrorCode = 1
                    result.ErrorMessage = String.Format(Language.resultMsg, aCustId, outPara.Item("p_RetMsg"))
                    result.ResultBoolean = True
                End If
                
            Else
                result.ResultBoolean = False
                result.ErrorCode = outPara("p_RetCode")

                If String.IsNullOrEmpty(outPara.Item("p_RetMsg")) Then
                    result.ErrorMessage = Language.noAnyInv
                Else
                    result.ErrorMessage = outPara("p_RetMsg")
                End If

            End If
            If blnAutoClose Then
                If result.ResultBoolean Then
                    trans.Commit()
                Else
                    trans.Rollback()
                End If

            End If
            'result.ResultBoolean = True
            Return result
        Catch ex As Exception
            If blnAutoClose Then
                trans.Rollback()
            End If
            result.ResultBoolean = False
            result.ErrorCode = -999
            result.ErrorMessage = ex.ToString

            Return result
        Finally
            If tb099 IsNot Nothing Then
                tb099.Dispose()
                tb099 = Nothing
            End If

        End Try

    End Function
    Private Function getInvoiceYearMonth(ByVal invDate As String) As String

        Dim trimInvdate As String = Replace(invDate, "/", "")
        Dim sL_Temp As String = trimInvdate.Substring(4, 2)
        If Integer.Parse(sL_Temp) Mod 2 = 0 Then sL_Temp = (Integer.Parse(sL_Temp) - 1).ToString
        If sL_Temp.Length < 2 Then sL_Temp = "0" & sL_Temp
        Return trimInvdate.Substring(0, 4) & sL_Temp
    
    End Function
#Region "IDisposable Support"
    Private disposedValue As Boolean ' 偵測多餘的呼叫

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
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
