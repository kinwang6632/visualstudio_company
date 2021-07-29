Imports System.Data.Common
Imports CableSoft.BLL.Utility
Public Class BillingAPI514
    Inherits BLLBasic
    Implements IDisposable, CableSoft.BLL.BillingAPI.IBillingAPI
    Private Language As New CableSoft.BLL.Language.SO61.BillingAPI514Language
    Private _DAL As New BillingAPI514DALMultiDB(Me.LoginInfo.Provider)
    Private Const FNewAccountTableName As String = "Account"
    'Private Const FOldProductTableName As String = "OldProduct"
    'Private Const FChangeProductTableName As String = "ChangeProduct"
    Private Const FOldProductTableName As String = "OldProduct"
    Private Const FChangeProductTableName As String = "ChangeProduct"
    Private Const inputAccount As String = "Account"
    Private Const FOldAccountTableName As String = "OldAccount"
    Private Const FDeclaredTableName As String = "Declared"
    Private Const FOldAch As String = "OldAch"
    Private Const FPKField As String = "MasterId"
    Private Const VoidBillTableName As String = "VoidBillNo"
    Private Account As Account = Nothing
    Private _AccountDAL As New AccountDAL(Me.LoginInfo.Provider)
    Private isACHBank As Boolean = False
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
    Protected Friend Property isACH As Boolean
      
        Set(ByVal Value As Boolean)
            isACHBank = Value
        End Set
        Get
            Return isACHBank
        End Get
    End Property


    Protected Friend Function insSO106FromInput(InData As System.Data.DataSet) As DataTable
        Dim compcode As Integer = Integer.Parse(InData.Tables("Main").Rows(0).Item("CompCode"))
        Dim tbSO106 As DataTable = Nothing
        Dim cmRefNo As Integer = 0
        Try
            With InData.Tables(inputAccount).Rows(0)
                tbSO106 = DAO.ExecQry(_DAL.getEmptySO106)

                Dim rwSO106 As DataRow = tbSO106.NewRow
                rwSO106.Item("AcceptEn") = .Item("AcceptEn")
                rwSO106.Item("AcceptName") = DAO.ExecSclr(_DAL.getEmpName, New Object() _
                                                            {.Item("AcceptEn"), compcode})
                rwSO106.Item("Proposer") = .Item("Proposer")
                rwSO106.Item("ID") = .Item("ID")
                rwSO106.Item("BankCode") = .Item("BankCode")
                rwSO106.Item("BankName") = DAO.ExecSclr(_DAL.getBankName, New Object() {Integer.Parse(.Item("BankCode"))})
                If Not DBNull.Value.Equals(.Item("CardCode")) Then
                    rwSO106.Item("CardCode") = .Item("CardCode")
                    rwSO106.Item("CardName") = DAO.ExecSclr(_DAL.getCardName, New Object() {Integer.Parse(.Item("CardCode"))})
                End If
                If Not DBNull.Value.Equals(.Item("StopYM")) Then
                    rwSO106("StopYM") = Integer.Parse(.Item("StopYM"))
                End If
                If Not DBNull.Value.Equals(.Item("AccountName")) Then
                    rwSO106.Item("AccountName") = .Item("AccountName")
                End If
                If Not DBNull.Value.Equals(.Item("AccountNameID")) Then
                    rwSO106.Item("AccountNameID") = .Item("AccountNameID")
                End If
                rwSO106.Item("PropDate") = Date.Parse(.Item("PropDate"))
                If Not DBNull.Value.Equals(.Item("SendDate")) Then
                    rwSO106.Item("SendDate") = Date.Parse(.Item("SendDate"))
                End If
                If Not DBNull.Value.Equals(.Item("SnactionDate")) Then
                    rwSO106.Item("SnactionDate") = Date.Parse(.Item("SnactionDate"))
                End If
                rwSO106.Item("StopFlag") = 0
                If Not DBNull.Value.Equals(.Item("StopDate")) Then
                    rwSO106.Item("StopDate") = Date.Parse(.Item("StopDate"))
                    rwSO106.Item("StopFlag") = 1
                End If
                If Not DBNull.Value.Equals(.Item("MediaCode")) Then
                    rwSO106.Item("MediaCode") = .Item("MediaCode")
                    rwSO106.Item("MediaName") = DAO.ExecSclr(_DAL.getMediaName, New Object() {Integer.Parse(.Item("MediaCode"))})
                End If
                If Not DBNull.Value.Equals(.Item("IntroID")) Then
                    rwSO106.Item("IntroID") = .Item("IntroID")
                    Dim intMediaRefNo As Integer = Integer.Parse(DAO.ExecSclr(_DAL.getMediaRefNo, New Object() {Integer.Parse(.Item("MediaCode"))}))
                    Dim aIntroID As Object = .Item("IntroID")
                    If intMediaRefNo <> 2 AndAlso intMediaRefNo <> 3 Then
                        If Not IsNumeric(.Item("IntroID")) Then
                            aIntroID = -1
                        End If
                    End If
                    Using o As New CableSoft.SO.BLL.Customer.IntroMedia.IntroMedia(Me.LoginInfo, Me.DAO)
                        Using t As DataTable = o.keyCodeSearch(intMediaRefNo, aIntroID)
                            If t.Rows.Count > 0 Then
                                rwSO106.Item("IntroName") = t.Rows(0).Item("Description")
                            End If
                        End Using
                    End Using

                End If
                rwSO106.Item("Note") = .Item("Note")
                rwSO106.Item("CompCode") = compcode
                rwSO106.Item("CustID") = Integer.Parse(.Item("CustID"))
                rwSO106.Item("CMCode") = Integer.Parse(.Item("CMCode"))
                rwSO106.Item("CMName") = DAO.ExecSclr(_DAL.getCMName, New Object() {Integer.Parse(.Item("CMCode"))})
                cmRefNo = DAO.ExecSclr(_DAL.getCD031RefNo, New Object() {Integer.Parse(.Item("CMCode"))})
                rwSO106.Item("UpdEn") = LoginInfo.EntryName
                rwSO106.Item("Alien") = .Item("Alien")
                rwSO106.Item("CVC2") = .Item("CVC2")
                rwSO106.Item("CitemStr") = .Item("CitemStr")
                rwSO106.Item("CitemStr2") = .Item("CitemStr2")
                rwSO106.Item("AddCitemAccount") = 0
                If Not DBNull.Value.Equals(.Item("AddCitemAccount")) Then
                    rwSO106.Item("AddCitemAccount") = Integer.Parse(.Item("AddCitemAccount"))
                End If
                rwSO106.Item("PTCode") = .Item("PTCode")
                rwSO106.Item("PTName") = DAO.ExecSclr(_DAL.getPTName, New Object() {Integer.Parse(.Item("PTCode"))})
                rwSO106.Item("ACHSN") = .Item("ACHSN")
                rwSO106.Item("ACHTNo") = .Item("ACHTNo")
                rwSO106.Item("AccountID") = .Item("AccountID")
                If cmRefNo = 2 OrElse cmRefNo = 4 OrElse cmRefNo = 5 Then
                    rwSO106.Item("AccountId") = .Item("AccountId")
                Else
                    Using Account As New Account(LoginInfo, Me.DAO)
                        Dim r As RIAResult = Account.GetOldVirtualAccount(Integer.Parse(.Item("CustId")), Integer.Parse(.Item("BankCode")))
                        rwSO106.Item("AccountId") = r.ResultXML
                    End Using
                End If
                tbSO106.Rows.Add(rwSO106)
            End With

        Catch ex As Exception
            Throw ex
        End Try
        Return tbSO106.Copy
    End Function
    Private Function CreateVoidBillTableName() As DataTable
        Dim tb As DataTable = New DataTable
        tb.Columns.Add("BillNo", GetType(String))
        tb.Columns.Add("Item", GetType(Integer))
        tb.TableName = VoidBillTableName
        Return tb
    End Function
    Friend Function isAchtNoOK(ByVal scrAchtNo As String, tbCD018 As DataTable, tbAchtNo As DataTable) As RIAResult
        scrAchtNo = scrAchtNo.Replace("'", "")
        Dim result As New RIAResult With {.ResultBoolean = False, .ErrorCode = -1, .ErrorMessage = Language.ACHTNONotSameBank}
        For Each strAchtNo As String In scrAchtNo.Split(",")
            If tbAchtNo.AsEnumerable.Count(Function(rw As DataRow)
                                               If rw.Item("ACHTNO") = strAchtNo AndAlso Integer.Parse(tbCD018.Rows(0).Item("ACHTYPE")) = Integer.Parse(rw.Item("ACHTYPE")) Then
                                                   result.ResultBoolean = True
                                                   result.ErrorCode = 0
                                                   result.ErrorMessage = Nothing
                                                   If String.IsNullOrEmpty(result.ResultXML) Then
                                                       result.ResultXML = String.Format("'{0}'", rw.Item("ACHTDESC"))
                                                   Else
                                                       result.ResultXML = String.Format("{0},'{1}'", result.ResultXML, rw.Item("ACHTDESC"))
                                                   End If
                                                   Return True
                                               Else
                                                   Return False
                                               End If

                                           End Function) <> 1 Then
                result.ResultBoolean = False
                result.ErrorCode = -1
                result.ErrorMessage = String.Format(Language.ACHTNONotSameBank2, strAchtNo)
                Return result
            End If
        Next
        Return result
    End Function

    Friend Function isDataOK(ByVal inData As DataSet, ByVal tbSO106 As DataTable, ByVal editMode As EditMode) As RIAResult
        Dim result As New RIAResult With {.ResultBoolean = False, .ErrorMessage = Nothing, .ErrorCode = -1}
        Dim errormsg As String = Nothing

        With inData.Tables(FNewAccountTableName).Rows(0)
            If inData.Tables(FNewAccountTableName).Columns.Contains("ID") Then
                If Not DBNull.Value.Equals(.Item("ID")) Then
                    If Not CableSoft.BLL.Utility.Utility.IDVerify(.Item("ID"), errormsg) Then
                        result.ErrorMessage = errormsg
                        Return result
                    End If
                End If
            End If
            If inData.Tables(FNewAccountTableName).Columns.Contains("AccountNameID") Then
                If Not DBNull.Value.Equals(.Item("AccountNameID")) Then
                    If Not CableSoft.BLL.Utility.Utility.IDVerify(.Item("AccountNameID"), errormsg) Then
                        result.ErrorMessage = Language.mustAccountName & errormsg
                        Return result
                    End If
                End If
            End If

            If isACHBank Then
                If inData.Tables(FNewAccountTableName).Columns.Contains("SendDate") Then
                    If Not DBNull.Value.Equals(.Item("SendDate")) Then
                        result.ErrorMessage = String.Format(Language.cannotSendDate, .Item("BankCode"))
                        Return result
                    End If
                End If
                If inData.Tables(FNewAccountTableName).Columns.Contains("SnactionDate") Then
                    If Not DBNull.Value.Equals(.Item("SnactionDate")) Then
                        result.ErrorMessage = String.Format(Language.cannotSnactionDate, .Item("BankCode"))
                        Return result
                    End If
                End If
                If inData.Tables(FNewAccountTableName).Columns.Contains("ACHTNo") Then
                    If DBNull.Value.Equals(.Item("ACHTNo")) Then
                        result.ErrorMessage = Language.noFoundACHTNO
                        Return result
                    End If
                Else
                    result.ErrorMessage = Language.noFoundACHTNO
                    Return result
                End If

            Else
                If inData.Tables(FNewAccountTableName).Columns.Contains("ACHTNo") Then
                    If Not DBNull.Value.Equals(.Item("ACHTNO")) Then
                        result.ErrorMessage = Language.cannotACHTNO
                        Return result
                    End If
                End If

            End If
            If Not DBNull.Value.Equals(.Item("StopDate")) AndAlso editMode = CableSoft.BLL.Utility.EditMode.Append Then
                result.ErrorMessage = Language.cannotStopDate
                Return result
            End If
            If editMode = CableSoft.BLL.Utility.EditMode.Append Then
                If Not DBNull.Value.Equals(.Item("DeAuthorize")) AndAlso Integer.Parse(.Item("DeAuthorize")) <> 0 Then
                    result.ErrorMessage = Language.cannotDeAuthorize
                    Return result
                End If
            End If

            If Not DBNull.Value.Equals(inData.Tables("Account").Rows(0).Item("IntroID")) Then
                If DBNull.Value.Equals(tbSO106.Rows(0).Item("IntroName")) AndAlso Not DBNull.Value.Equals(tbSO106.Rows(0).Item("MediaCode")) Then
                    result.ErrorMessage = Language.noFoundIntroName
                    result.ErrorCode = -166
                    Return result
                End If
            End If


        End With
        If String.IsNullOrEmpty(result.ErrorMessage) Then
            result.ResultBoolean = True
            result.ErrorCode = 0
        End If
        Return result
    End Function
    Public Function Execute(SeqNo As Integer, InData As System.Data.DataSet) As CableSoft.BLL.Utility.RIAResult Implements CableSoft.BLL.BillingAPI.IBillingAPI.Execute
        Dim cn As DbConnection = DAO.GetConn()
        Dim trans As DbTransaction = Nothing
        Dim blnAutoClose As Boolean = False
        Dim result As New RIAResult
        Dim StartPost As Boolean = False
        Dim tbCD018 As DataTable = Nothing
        Dim tbActhNO As DataTable = Nothing
        Dim tbOupt As DataTable = New DataTable("INV")
        Dim dsOutput As New DataSet
        tbOupt.Columns.Add("INVSEQNO", GetType(String))
        tbOupt.Columns.Add("ACHCustId", GetType(String))
        tbOupt.Columns.Add("AccountID", GetType(String))
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
        Account = New Account(Me.LoginInfo, DAO)
        Account.setFlow = False

        StartPost = Integer.Parse(DAO.ExecSclr(_AccountDAL.QueryStartPost, New Object() {LoginInfo.CompCode})) = 1
        '#8706
        Me.LoginInfo.EntryId = InData.Tables("Account").Rows(0).Item("AcceptEn")
        Using EmpName As DataTable = DAO.ExecQry(_DAL.getEmpName(), New Object() {InData.Tables("Account").Rows(0).Item("AcceptEn"), Me.LoginInfo.CompCode})
            Me.LoginInfo.EntryId = InData.Tables("Account").Rows(0).Item("AcceptEn")
            Me.LoginInfo.EntryName = EmpName.Rows(0).Item("EmpName")
        End Using

        Try

            With InData.Tables(inputAccount).Rows(0)
                If Not DBNull.Value.Equals(.Item("IntroID")) AndAlso DBNull.Value.Equals(.Item("MediaCode")) Then
                    result.ErrorCode = -1
                    result.ErrorMessage = Language.mustMediaCode
                    result.ResultBoolean = False
                    Return result
                End If
                Using dsSO106 As New DataSet
                    dsSO106.Tables.Add(InData.Tables("INV").Copy)
                    dsSO106.Tables.Add(CreateVoidBillTableName.Copy)
                    Using tbSO106 As DataTable = insSO106FromInput(InData)
                        tbSO106.TableName = FNewAccountTableName
                        dsSO106.Tables.Add(tbSO106.Copy)
                    End Using
                    If Account.IsACHBank(InData.Tables(FNewAccountTableName).Rows(0).Item("BankCode"), StartPost).ResultBoolean Then
                        isACHBank = True

                        tbCD018 = DAO.ExecQry(_AccountDAL.GetBankCodeByCode(StartPost), _
                                         New Object() {Integer.Parse(InData.Tables(FNewAccountTableName).Rows(0).Item("BankCode"))})
                        tbActhNO = DAO.ExecQry(_AccountDAL.GetACHTNo(StartPost))
                        If DBNull.Value.Equals(InData.Tables(FNewAccountTableName).Rows(0).Item("ACHTNO")) Then
                            result.ResultBoolean = False
                            result.ErrorCode = -706
                            result.ErrorMessage = Language.musAchtNo
                            Return result
                        End If
                        result = isAchtNoOK(InData.Tables(FNewAccountTableName).Rows(0).Item("ACHTNO"), tbCD018, tbActhNO)
                        If result.ResultBoolean Then
                            dsSO106.Tables(FNewAccountTableName).Rows(0).Item("ACHTDESC") = result.ResultXML
                            dsSO106.Tables(FNewAccountTableName).AcceptChanges()
                        Else
                            Return result
                        End If


                    End If
                    result = isDataOK(InData, dsSO106.Tables(FNewAccountTableName), EditMode.Append)
                    If Not result.ResultBoolean Then Return result


                    Using tbChangeProduct As New DataTable
                        tbChangeProduct.Columns.Add("CustId", GetType(Integer))
                        tbChangeProduct.Columns.Add("ServiceId", GetType(String))
                        Dim r As DataRow = tbChangeProduct.NewRow
                        r.Item("CustId") = dsSO106.Tables(FNewAccountTableName).Rows(0).Item("CustId")
                        r.Item("ServiceId") = dsSO106.Tables(FNewAccountTableName).Rows(0).Item("CitemStr")
                        tbChangeProduct.Rows.Add(r)
                        tbChangeProduct.TableName = FChangeProductTableName
                        dsSO106.Tables.Add(tbChangeProduct)
                    End Using


                    result = Account.ChkDataOk(EditMode.Append, dsSO106)
                    If result.ResultBoolean Then
                        Using dsResult As DataSet = Account.SaveNewData(EditMode.Append, dsSO106)
                            Dim rwOutput As DataRow = tbOupt.NewRow
                            rwOutput.Item("INVSEQNO") = dsResult.Tables(FNewAccountTableName).Rows(0).Item("INVSEQNO")
                            If Not DBNull.Value.Equals(dsResult.Tables(FNewAccountTableName).Rows(0).Item("ACHCustId")) Then
                                rwOutput.Item("ACHCustId") = dsResult.Tables(FNewAccountTableName).Rows(0).Item("ACHCustId")
                            End If
                            If DBNull.Value.Equals(InData.Tables(FNewAccountTableName).Rows(0).Item("AccountId")) Then
                                rwOutput.Item("AccountID") = dsResult.Tables(FNewAccountTableName).Rows(0).Item("AccountID")
                            End If
                            tbOupt.Rows.Add(rwOutput)
                            If DBNull.Value.Equals(tbOupt.Rows(0).Item("ACHCustId")) Then
                                tbOupt.Columns.Remove("ACHCustId")
                            End If
                            If DBNull.Value.Equals(tbOupt.Rows(0).Item("AccountID")) Then
                                tbOupt.Columns.Remove("AccountID")
                            End If
                            dsOutput.Tables.Add(tbOupt)
                            result.ResultDataSet = dsOutput.Copy
                        End Using

                    End If

                End Using
            End With

            Return result
        Catch ex As Exception
            result.ResultBoolean = False
            result.ErrorCode = -99
            result.ErrorMessage = ex.ToString
        Finally
            If tbCD018 IsNot Nothing Then
                tbCD018.Dispose()
                tbCD018 = Nothing
            End If
            If tbActhNO IsNot Nothing Then
                tbActhNO.Dispose()
                tbActhNO = Nothing
            End If
            If tbOupt IsNot Nothing Then
                tbOupt.Dispose()
                tbOupt = Nothing
            End If
            If dsOutput IsNot Nothing Then
                dsOutput.Dispose()
                dsOutput = Nothing
            End If
        End Try

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
            If Account IsNot Nothing Then
                Account.Dispose()
                Account = Nothing
            End If
            If _AccountDAL IsNot Nothing Then
                _AccountDAL.Dispose()
                _AccountDAL = Nothing
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
