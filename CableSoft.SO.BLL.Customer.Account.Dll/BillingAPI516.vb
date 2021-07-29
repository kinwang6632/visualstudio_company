Imports System.Data.Common
Imports CableSoft.BLL.Utility
Public Class BillingAPI516
    Inherits BLLBasic
    Implements IDisposable, CableSoft.BLL.BillingAPI.IBillingAPI
    Private Language As New CableSoft.BLL.Language.SO61.BillingAPI516Language
    Private _DAL As New BillingAPI516DALMultiDB(Me.LoginInfo.Provider)
    Private _API514 As BillingAPI514 = Nothing
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
    Private oldSendDate As String = Nothing
    Private oldAccountId As String = Nothing
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
   
    Protected Friend Function insSO106FromInput(InData As System.Data.DataSet) As DataTable
        Dim compcode As Integer = Integer.Parse(InData.Tables("Main").Rows(0).Item("CompCode"))
        Dim tbSO106 As DataTable = Nothing
        Dim cmRefNo As Integer = 0
        Dim _514DAL As New BillingAPI514DAL(Me.LoginInfo.Provider)
        Try
            With InData.Tables(inputAccount).Rows(0)
                tbSO106 = DAO.ExecQry(_DAL.getSO106ByMasterId, New Object() {Integer.Parse(.Item("MasterId"))})
                oldAccountId = tbSO106.Rows(0).Item("AccountID")
                oldSendDate = Nothing
                If Not DBNull.Value.Equals(tbSO106.Rows(0).Item("SendDate")) Then
                    oldSendDate = "X"
                End If

                Dim rwSO106 As DataRow = tbSO106.Rows(0)
                'rwSO106.Item("AcceptEn") = .Item("AcceptEn")
                'rwSO106.Item("AcceptName") = DAO.ExecSclr(_514DAL.getEmpName, New Object() _
                '                                            {.Item("AcceptEn"), compcode})
                'rwSO106.Item("Proposer") = .Item("Proposer")
                'rwSO106.Item("ID") = .Item("ID")
                rwSO106.Item("BankCode") = .Item("BankCode")
                rwSO106.Item("BankName") = DAO.ExecSclr(_514DAL.getBankName, New Object() {Integer.Parse(.Item("BankCode"))})
                If Not DBNull.Value.Equals(.Item("CardCode")) Then
                    rwSO106.Item("CardCode") = .Item("CardCode")
                    rwSO106.Item("CardName") = DAO.ExecSclr(_514DAL.getCardName, New Object() {Integer.Parse(.Item("CardCode"))})
                Else
                    rwSO106.Item("CardCode") = DBNull.Value
                    rwSO106.Item("CardName") = DBNull.Value
                End If
                rwSO106.Item("AccountID") = .Item("AccountID")

                If Not DBNull.Value.Equals(.Item("CVC2")) Then
                    rwSO106.Item("CVC2") = .Item("CVC2")
                Else
                    rwSO106.Item("CVC2") = DBNull.Value
                End If

                If Not DBNull.Value.Equals(.Item("StopYM")) Then
                    rwSO106("StopYM") = Integer.Parse(.Item("StopYM"))
                Else
                    rwSO106("StopYM") = DBNull.Value
                End If

                If Not DBNull.Value.Equals(.Item("MediaCode")) Then
                    rwSO106.Item("MediaCode") = .Item("MediaCode")
                    rwSO106.Item("MediaName") = DAO.ExecSclr(_514DAL.getMediaName, New Object() {Integer.Parse(.Item("MediaCode"))})
                Else
                    rwSO106.Item("MediaCode") = DBNull.Value
                    rwSO106.Item("MediaName") = DBNull.Value
                End If
                If Not DBNull.Value.Equals(.Item("IntroID")) Then
                    rwSO106.Item("IntroID") = .Item("IntroID")
                    Dim intMediaRefNo As Integer = Integer.Parse(DAO.ExecSclr(_514DAL.getMediaRefNo, New Object() {Integer.Parse(.Item("MediaCode"))}))
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
                Else
                    rwSO106.Item("IntroID") = DBNull.Value
                End If
                rwSO106.Item("UpdEn") = .Item("UpdEn")
                'rwSO106.Item("UpdEn") = Me.LoginInfo.EntryName
                Me.LoginInfo.EntryName = .Item("UpdEn")
                rwSO106.Item("StopFlag") = 0
                If Not DBNull.Value.Equals(.Item("StopDate")) Then
                    rwSO106.Item("StopDate") = Date.Parse(.Item("StopDate"))
                    rwSO106.Item("StopFlag") = 1
                Else
                    rwSO106.Item("StopDate") = DBNull.Value
                End If
                If Not DBNull.Value.Equals(.Item("DeAuthorize")) Then
                    rwSO106.Item("DeAuthorize") = .Item("DeAuthorize")
                Else
                    rwSO106.Item("DeAuthorize") = DBNull.Value
                End If
                If Not DBNull.Value.Equals(.Item("AuthorizeStopDate")) Then
                    rwSO106("AuthorizeStopDate") = Date.Parse(.Item("AuthorizeStopDate"))
                Else
                    rwSO106("AuthorizeStopDate") = DBNull.Value
                End If
                If Not DBNull.Value.Equals(.Item("ACHSN")) Then
                    rwSO106.Item("ACHSN") = .Item("ACHSN")
                Else
                    rwSO106.Item("ACHSN") = DBNull.Value
                End If
                If Not DBNull.Value.Equals(.Item("ACHTNo")) Then
                    rwSO106.Item("ACHTNo") = .Item("ACHTNo")
                Else
                    rwSO106.Item("ACHTNo") = DBNull.Value
                End If

                If Not DBNull.Value.Equals(.Item("CitemStr")) Then
                    rwSO106.Item("CitemStr") = .Item("CitemStr")
                Else
                    rwSO106.Item("CitemStr") = DBNull.Value
                End If
                If Not DBNull.Value.Equals(.Item("CitemStr2")) Then
                    rwSO106.Item("CitemStr2") = .Item("CitemStr2")
                Else
                    rwSO106.Item("CitemStr2") = DBNull.Value
                End If
                rwSO106.Item("AddCitemAccount") = 0
                If Not DBNull.Value.Equals(.Item("AddCitemAccount")) Then
                    rwSO106.Item("AddCitemAccount") = Integer.Parse(.Item("AddCitemAccount"))

                End If
                If Not DBNull.Value.Equals(.Item("Note")) Then
                    rwSO106.Item("Note") = .Item("Note")
                Else
                    rwSO106.Item("Note") = DBNull.Value
                End If


                rwSO106.Item("CMCode") = Integer.Parse(.Item("CMCode"))
                rwSO106.Item("CMName") = DAO.ExecSclr(_514DAL.getCMName, New Object() {Integer.Parse(.Item("CMCode"))})
                'cmRefNo = DAO.ExecSclr(_514DAL.getCD031RefNo, New Object() {Integer.Parse(.Item("CMCode"))})

                rwSO106.Item("PTCode") = Integer.Parse(.Item("PTCode"))
                rwSO106.Item("PTName") = DAO.ExecSclr(_514DAL.getPTName, New Object() {Integer.Parse(.Item("PTCode"))})


            End With

        Catch ex As Exception
            Throw ex
        End Try
        Return tbSO106.Copy
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
        _API514 = New BillingAPI514(Me.LoginInfo, DAO)
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


        Try

            With InData.Tables(inputAccount).Rows(0)
                If Not DBNull.Value.Equals(.Item("IntroID")) AndAlso DBNull.Value.Equals(.Item("MediaCode")) Then
                    result.ErrorCode = -1
                    result.ErrorMessage = Language.mustMediaCode
                    result.ResultBoolean = False
                    Return result
                End If
                Using dsSO106 As New DataSet                    
                    dsSO106.Tables.Add(CreateVoidBillTableName.Copy)

                    Using tbSO106 As DataTable = insSO106FromInput(InData)
                        tbSO106.TableName = FNewAccountTableName
                        dsSO106.Tables.Add(tbSO106.Copy)
                    End Using
                    dsSO106.Tables.Add(CreateInvTable(dsSO106.Tables(FNewAccountTableName).Rows(0).Item("Invseqno")).Copy)
                    dsSO106.Tables.Add(CreateOldAchTable(dsSO106.Tables(FNewAccountTableName).Rows(0).Item("Masterid")).Copy)
                    If dsSO106.Tables("INV").Rows.Count = 0 Then
                        result.ErrorCode = -1
                        result.ErrorMessage = Language.noFoundInv
                        result.ResultBoolean = False
                        Return result
                    End If
                    If Account.IsACHBank(InData.Tables(FNewAccountTableName).Rows(0).Item("BankCode"), StartPost).ResultBoolean Then
                        isACHBank = True
                        _API514.isACH = isACHBank
                        tbCD018 = DAO.ExecQry(_AccountDAL.GetBankCodeByCode(StartPost), _
                                         New Object() {Integer.Parse(InData.Tables(FNewAccountTableName).Rows(0).Item("BankCode"))})
                        tbActhNO = DAO.ExecQry(_AccountDAL.GetACHTNo(StartPost))
                        If DBNull.Value.Equals(InData.Tables(FNewAccountTableName).Rows(0).Item("ACHTNO")) Then
                            result.ResultBoolean = False
                            result.ErrorCode = -706
                            result.ErrorMessage = Language.musAchtNo
                            Return result
                        End If
                        result = _API514.isAchtNoOK(InData.Tables(FNewAccountTableName).Rows(0).Item("ACHTNO"), tbCD018, tbActhNO)
                        If result.ResultBoolean Then
                            dsSO106.Tables(FNewAccountTableName).Rows(0).Item("ACHTDESC") = result.ResultXML
                            dsSO106.Tables(FNewAccountTableName).AcceptChanges()
                        Else
                            Return result
                        End If


                    End If

                    result = _API514.isDataOK(InData, dsSO106.Tables(FNewAccountTableName), EditMode.Edit)
                    If result.ResultBoolean Then
                        If (isACHBank) AndAlso (Not String.IsNullOrEmpty(oldSendDate)) Then
                            If dsSO106.Tables(FNewAccountTableName).Rows(0).Item("AccountId") <> oldAccountId Then
                                result.ResultBoolean = False
                                result.ErrorMessage = Language.cannotModiACHAccount
                                result.ErrorCode = -709

                            End If
                        End If
                    End If
                   
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


                    result = Account.ChkDataOk(EditMode.Edit, dsSO106)

                    If result.ResultBoolean Then
                        Using dsResult As DataSet = Account.SaveNewData(EditMode.Edit, dsSO106)
                            'Dim rwOutput As DataRow = tbOupt.NewRow
                            'rwOutput.Item("INVSEQNO") = dsResult.Tables(FNewAccountTableName).Rows(0).Item("INVSEQNO")
                            'If Not DBNull.Value.Equals(dsResult.Tables(FNewAccountTableName).Rows(0).Item("ACHCustId")) Then
                            '    rwOutput.Item("ACHCustId") = dsResult.Tables(FNewAccountTableName).Rows(0).Item("ACHCustId")
                            'End If
                            'If DBNull.Value.Equals(InData.Tables(FNewAccountTableName).Rows(0).Item("AccountId")) Then
                            '    rwOutput.Item("AccountID") = dsResult.Tables(FNewAccountTableName).Rows(0).Item("AccountID")
                            'End If
                            'tbOupt.Rows.Add(rwOutput)
                            'If DBNull.Value.Equals(tbOupt.Rows(0).Item("ACHCustId")) Then
                            '    tbOupt.Columns.Remove("ACHCustId")
                            'End If
                            'If DBNull.Value.Equals(tbOupt.Rows(0).Item("AccountID")) Then
                            '    tbOupt.Columns.Remove("AccountID")
                            'End If
                            'dsOutput.Tables.Add(tbOupt)
                            'result.ResultDataSet = dsResult
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
            If _API514 IsNot Nothing Then
                _API514.Dispose()
                _API514 = Nothing
            End If
            If _AccountDAL IsNot Nothing Then
                _AccountDAL.Dispose()
                _AccountDAL = Nothing
            End If
            If _DAL IsNot Nothing Then
                _DAL.Dispose()
                _DAL = Nothing
            End If
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
    Private Function CreateInvTable(ByVal invseqno As Integer) As DataTable
        Using tb As DataTable = DAO.ExecQry(_DAL.getINV138ByInvseqno, New Object() {invseqno})
            tb.TableName = "INV"
            Return tb.Copy
        End Using


    End Function
    Private Function CreateOldAchTable(ByVal Masterid As Integer) As DataTable

        Using tb = DAO.ExecQry(_DAL.getOldACHByMasterid, New Object() {Masterid})
            tb.TableName = FOldAch
            Return tb.Copy
        End Using

    End Function
    Private Function CreateVoidBillTableName() As DataTable
        Dim tb As DataTable = New DataTable
        tb.Columns.Add("BillNo", GetType(String))
        tb.Columns.Add("Item", GetType(Integer))
        tb.TableName = VoidBillTableName
        Return tb
    End Function
#Region "IDisposable Support"
    Private disposedValue As Boolean ' 偵測多餘的呼叫

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
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
