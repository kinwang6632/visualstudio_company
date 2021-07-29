Imports System.Data.Common
Imports CableSoft.BLL.BillingAPI
Imports CableSoft.BLL.Utility
Public Class BillingAPI602
    Inherits BLLBasic
    Implements IDisposable, CableSoft.BLL.BillingAPI.IBillingAPI
    Private _DAL As New BillingAPI602DAL(Me.LoginInfo.Provider)
    Private Language As New CableSoft.BLL.Language.SO61.BillingAPI602Language
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
    Public Function Execute(SeqNo As Integer, InData As DataSet) As RIAResult Implements IBillingAPI.Execute
        Dim aCusComp As Object = InData.Tables("Main").Rows(0).Item("CusComp")
        Dim aCusOwner As Object = InData.Tables("Main").Rows(0).Item("CusOwner")
        Dim aLinkToMIS As Boolean = False
        Dim aInvID As Object = InData.Tables("Inv").Rows(0).Item("InvID")
        Dim aCompCode As Integer = InData.Tables("Main").Rows(0).Item("Compcode")
        Dim aPaperDate As Date = Date.Parse(InData.Tables("Inv").Rows(0).Item("PaperDate"))
        Dim aInvAmount As Integer = Integer.Parse(InData.Tables("Inv").Rows(0).Item("InvAmount"))
        Dim aCaller As String = InData.Tables("Main").Rows(0).Item("Caller").ToString
        Dim tbMainInv007 As DataTable = Nothing
        Dim result As New RIAResult With {.ResultBoolean = False, .ErrorCode = -1, .ErrorMessage = "NO"}
        Dim tbSO034 As DataTable = Nothing
        Dim aYearMonth As String = Nothing
        Dim aAllowanceTotal As Integer = 0
        Dim aUpdlimtDate As Integer = -1
        Dim aTaxAmount As Integer = 0
        Dim aSaleAmount As Integer = 0
        Dim aTaxType As Integer = 0
        Dim aAllowanceNo As String = Nothing
        Dim cn As DbConnection = DAO.GetConn()
        Dim trans As DbTransaction = Nothing
        Dim blnAutoClose As Boolean = False
        Dim updTime As DateTime = DateTime.Now
        Dim aBillNoItem As String = "'X'"
        Dim aNote As String = Nothing
        '#8706
        aCaller = InData.Tables("Inv").Rows(0).Item("Upden")
        Me.LoginInfo.EntryName = aCaller
        'aCaller = Me.LoginInfo.EntryName
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
        Try
            aUpdlimtDate = DAO.ExecSclr(_DAL.QueryUpdlimtDate, New Object() {aCompCode})
            If aPaperDate.Month Mod 2 = 0 Then
                aYearMonth = aPaperDate.Year.ToString & Right("00" & (aPaperDate.Month - 1).ToString, 2)
            Else
                aYearMonth = aPaperDate.Year & Right("00" & aPaperDate.Month.ToString, 2)
            End If
            If Not DBNull.Value.Equals(aCusComp) AndAlso Not DBNull.Value.Equals(aCusOwner) Then
                aLinkToMIS = True
                tbSO034 = DAO.ExecQry(_DAL.QuerySO034(aCusOwner), New Object() {aInvID, aCusComp})
            End If
           
            aAllowanceTotal = aInvAmount + Integer.Parse(DAO.ExecSclr(_DAL.QuerySumAmt, New Object() {aInvID, aCompCode}))
            tbMainInv007 = DAO.ExecQry(_DAL.QueryInv007, New Object() {aInvID, aCompCode})
            aTaxType = Integer.Parse(tbMainInv007.Rows(0).Item("TaxType").ToString())
            CaculateTax(aTaxType, aInvAmount, aTaxAmount, aSaleAmount)
            If aInvAmount > Integer.Parse(tbMainInv007.Rows(0).Item("INVAMOUNT")) Then
                result.ResultBoolean = False
                result.ErrorCode = -632
                result.ErrorMessage = Language.DiscountLargerInv
                Return result
            End If

            If aAllowanceTotal > Integer.Parse(tbMainInv007.Rows(0).Item("INVAMOUNT")) Then
                result.ResultBoolean = False
                result.ErrorCode = -632
                result.ErrorMessage = Language.SumDiscountLargerInv
                Return result
            End If
            If Integer.Parse(tbMainInv007.Rows(0).Item("INVAMOUNT")) <> aAllowanceTotal Then aUpdlimtDate = -1
            If aUpdlimtDate > 0 AndAlso 1 = 0 Then
                If Not CanAllowance(Now, Date.Parse(tbMainInv007.Rows(0).Item("INVDATE")), aUpdlimtDate) Then
                    result.ResultBoolean = False
                    result.ErrorCode = -633
                    result.ErrorMessage = Language.onlyCanDrop
                    Return result
                End If
            End If
            If aLinkToMIS Then
                If tbSO034.Rows.Count = 0 Then
                    result.ResultBoolean = False
                    result.ErrorCode = -637
                    result.ErrorMessage = Language.noFoundInvId
                    Return result

                End If
                Dim aAmt034 As Integer = tbSO034.AsEnumerable.Sum(Function(rw As DataRow)
                                                                      Return rw.Item("REALAMT")
                                                                  End Function)
                If Math.Abs(aAmt034) <> aInvAmount Then
                    result.ResultBoolean = False
                    result.ErrorCode = -638
                    result.ErrorMessage = Language.notSameCust
                    Return result
                End If

                If aAmt034 > 0 Then
                    result.ResultBoolean = False
                    result.ErrorCode = -635
                    result.ErrorMessage = Language.noLargerZero
                    Return result
                End If

            End If
            If Integer.Parse(DAO.ExecSclr(_DAL.QueryIsDataLocked, New Object() {aCompCode, aYearMonth})) > 0 Then
                result.ResultBoolean = False
                result.ErrorCode = -1
                result.ErrorMessage = Language.beLock
                Return result
            End If
            Dim isDual As Boolean = False
            Dim o As Object = DAO.ExecSclr(_DAL.QueryDualInv014, New Object() {aInvID, New DateTime(aPaperDate.Year, aPaperDate.Month, aPaperDate.Day), aCompCode})
            If o IsNot Nothing Then
                aAllowanceNo = o.ToString()
            End If
            If Not String.IsNullOrEmpty(aAllowanceNo) Then
                isDual = True
                CaculateTax(aTaxType, aAllowanceTotal, aTaxAmount, aSaleAmount)
            Else
                'aAllowanceNo = DAO.ExecSclr(_DAL.GetAllowanceNo)
            End If
            aAllowanceNo = DAO.ExecSclr(_DAL.GetAllowanceNo)
            isDual = False
            CaculateTax(aTaxType, aInvAmount, aTaxAmount, aSaleAmount)
            'insert or update master
            If Not isDual Then
                DAO.ExecQry(_DAL.InsInv014, New Object() {aCompCode, _
                                                      tbMainInv007.Rows(0).Item("CustId"), New Date(aPaperDate.Year, aPaperDate.Month, aPaperDate.Day), tbMainInv007.Rows(0).Item("BUSINESSID"), _
                                                      aYearMonth, aInvID, 0, _
                                                      tbMainInv007.Rows(0).Item("INVFORMAT"), tbMainInv007.Rows(0).Item("INVDATE"), aTaxType, _
                                                      aSaleAmount, aTaxAmount, aInvAmount, _
                                                      updTime, aCaller, aAllowanceNo, _
                                                      0, aAllowanceNo, 0})
            Else
                DAO.ExecQry(_DAL.UpdInv014, New Object() {aAllowanceTotal, updTime, aCaller, aSaleAmount, _
                                                          aTaxAmount, aInvID, aAllowanceNo, aCompCode})

            End If
            'insert or update detail
            If Not isDual Then
                If Not aLinkToMIS Then
                    DAO.ExecQry(_DAL.insInv014A, New Object() {DBNull.Value, aAllowanceNo, 0, DBNull.Value, _
                                                               DBNull.Value, DBNull.Value, DBNull.Value, _
                                                               aInvAmount, updTime, aCaller, _
                                                               aAllowanceNo})
                Else
                    'For Each row As DataRow In tbSO034.Rows
                    '    DAO.ExecQry(_DAL.insInv014A, New Object() {row("ServiceType"), row("BillNo"), row("Item"), row("CitemCode"), _
                    '                                           row("CitemName"), row("PTCODE"), row("PTName"), _
                    '                                           Integer.Parse(row("RealAmt")) * -1, updTime, aCaller, _
                    '                                           aAllowanceNo})                       
                    'Next
                End If
            Else
                If Not aLinkToMIS Then
                    DAO.ExecQry(_DAL.updInv014A, New Object() {aAllowanceTotal, updTime, aCaller, aAllowanceNo})                                      
                Else                    
                    Using tbINV014A As DataTable = DAO.ExecQry(_DAL.QueryInv014A, New Object() {aAllowanceNo})
                        For Each row As DataRow In tbINV014A.Rows
                            aBillNoItem = String.Format("{0},'{1}'", aBillNoItem, row("PaperNo") & row("Seq"))
                        Next
                        tbINV014A.Dispose()
                    End Using
                    If tbSO034 IsNot Nothing Then
                        tbSO034.Dispose()
                        tbSO034 = Nothing
                    End If
                    tbSO034 = DAO.ExecQry(_DAL.QueryExcludeBill(aCusOwner, aBillNoItem), New Object() {aInvID, aCusComp})
                End If
            End If

            If tbSO034 IsNot Nothing Then
                For Each row As DataRow In tbSO034.Rows
                    aNote = Nothing
                    DAO.ExecQry(_DAL.insInv014A, New Object() {row("ServiceType"), row("BillNo"), row("Item"), row("CitemCode"), _
                                                           row("CitemName"), row("PTCODE"), row("PTName"), _
                                                           Integer.Parse(row("RealAmt")) * -1, updTime, aCaller, _
                                                           aAllowanceNo})
                    If Not DBNull.Value.Equals(row.Item("Note")) Then
                        aNote = row.Item("Note") & "," & String.Format(Language.noteMsg, aAllowanceNo)
                    Else
                        aNote = String.Format(Language.noteMsg, aAllowanceNo)
                    End If
                    DAO.ExecNqry(_DAL.updSO034(aCusOwner), New Object() {updTime, aNote, _
                                                                         row.Item("BillNo"), row.Item("Item"), _
                                                                         aCusComp})
                 
                Next
            End If
            

          
            If blnAutoClose Then
                If result.ResultBoolean Then
                    trans.Commit()
                Else
                    trans.Rollback()
                End If
            End If
            result.ResultBoolean = True
            result.ErrorCode = 1
            'result.ErrorMessage = String.Format(Language.resultMsg, aInvID)
            result.ErrorCode = 1
            result.ErrorMessage = String.Format(Language.retMsg, aInvID, aAllowanceNo)
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
            If tbMainInv007 IsNot Nothing Then
                tbMainInv007.Dispose()
                tbMainInv007 = Nothing
            End If
            If tbSO034 IsNot Nothing Then
                tbSO034.Dispose()
                tbSO034 = Nothing
            End If
        End Try
    End Function
    Private Sub CaculateTax(ByVal aTaxType As Integer, ByVal aInvAmount As Integer, ByRef aTaxAmount As Integer, ByRef aSaleAmount As Integer)
        aTaxAmount = 0
        aSaleAmount = 0
        If aTaxType = 1 Then

            aSaleAmount = Math.Round((aInvAmount / 1.05),
                                            0, MidpointRounding.AwayFromZero)
            aTaxAmount = aInvAmount - aSaleAmount
        Else
            aTaxAmount = 0
            aSaleAmount = aInvAmount
        End If
    End Sub
    Private Function CanAllowance(ByVal aClaimYearMonth As Date, ByVal aInvDate As Date, ByVal aUpdlimtDate As Integer) As Boolean
        Dim aAppDate As New Date(aClaimYearMonth.Year, aClaimYearMonth.Month, aUpdlimtDate)
        Dim aNowDate As New Date(Now.Year, Now.Month, Now.Day)
        Dim result As Boolean = False
        Select Case aInvDate.Month
            Case 1, 3, 5, 7, 9, 11
                'aInvDate = aInvDate.AddMonths(2)
                aInvDate = CableSoft.BLL.Utility.Utility.GetAddMonths(aInvDate, 2, Me.DAO)
            Case Else
        End Select
        Select Case aNowDate.Month
            Case 1, 3, 5, 7, 9, 11
                If aNowDate.Day > aUpdlimtDate Then
                    'aNowDate = aNowDate.AddMonths(2)
                    aNowDate = CableSoft.BLL.Utility.Utility.GetAddMonths(aNowDate, 2, Me.DAO)
                End If
        End Select
        If Date.Compare(aNowDate, aInvDate) > 0 Then
            result = True
        Else
            result = False
        End If
        If result = False AndAlso Date.Compare(aAppDate, aInvDate) > 0 Then
            result = True
        End If
        Return result
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
