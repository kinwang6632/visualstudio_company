Imports System.Data.Common
Imports CableSoft.BLL.Utility
'Imports Lang = CableSoft.SO.BLL.Customer.Account.AccountLanguage
'Imports System.Windows.Forms
Public Class Account
    Inherits BLLBasic
    Implements IDisposable

    Private _DAL As New AccountDALMultiDB(Me.LoginInfo.Provider)
    Private Const FNewAccountTableName As String = "Account"
    'Private Const FOldProductTableName As String = "OldProduct"
    'Private Const FChangeProductTableName As String = "ChangeProduct"
    Private Const FOldProductTableName As String = "OldProduct"
    Private Const FChangeProductTableName As String = "ChangeProduct"

    Private Const FOldAccountTableName As String = "OldAccount"
    Private Const FDeclaredTableName As String = "Declared"
    Private Const FOldAch As String = "OldAch"
    Private Const FPKField As String = "MasterId"
    Private Const VoidBillTableName As String = "VoidBillNo"
    Private FNowDate As Date = Date.Now
    Private Language As New CableSoft.BLL.Language.SO61.AccountLanguage
    Private newFlow As Boolean = True
    Private SO138_InvSeqNo As Object
    Public Enum AuthStatus
        Auth = 0
        Cancel = 1
    End Enum
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

    Public Function QueryAccountDetail(ByVal MasterId As Int32) As DataTable
        Return DAO.ExecQry(_DAL.QueryAccountDetail, New Object() {MasterId})
    End Function
    Public Function QueryReLoadData(ByVal MasterId As Integer) As DataSet
        Return QueryReLoadData(MasterId, "", True)
    End Function
    Public Property setFlow As Boolean
        Set(value As Boolean)
            newFlow = value
        End Set
        Get
            Return newFlow
        End Get
    End Property
    Public Function QueryOldReLoadData(ByVal MasterId As Integer) As DataSet
        Dim dsReturn As New DataSet
        Dim tbAccount As DataTable = DAO.ExecQry(_DAL.QueryAccount(), New Object() {MasterId})
        tbAccount.TableName = "Account"
        dsReturn.Tables.Add(tbAccount.Copy)
        Return dsReturn.Copy
    End Function

    Public Function QueryReLoadData(ByVal MasterId As Integer, ByVal SEQNO As String, ByVal filterCustId As Boolean) As DataSet
        'DAO.AutoCloseConn = False
        Dim dsReturn As New DataSet
        Dim CustId As Integer = -1
        Dim strProServiceId As String = "-1"
        Try
            Dim tbAccount As DataTable = DAO.ExecQry(_DAL.QueryAccount(), New Object() {MasterId}).Copy
            If filterCustId Then
                If (tbAccount IsNot Nothing) AndAlso (tbAccount.Rows.Count > 0) Then
                    CustId = Integer.Parse(tbAccount.Rows(0).Item("CustId"))
                    If (DBNull.Value.Equals(tbAccount.Rows(0).Item("SnactionDate"))) AndAlso
                              (Not DBNull.Value.Equals(tbAccount.Rows(0).Item("ProServiceID"))) Then
                        strProServiceId = tbAccount.Rows(0).Item("ProServiceID")
                    End If

                End If
            End If
            tbAccount.TableName = "Account"
            dsReturn.Tables.Add(tbAccount)

            Dim tbChooseProduct As DataTable = Nothing
            If filterCustId Then
                tbChooseProduct = QueryChooseProduct(CustId, MasterId).Copy()
            Else
                tbChooseProduct = QueryNewChooseProduct(MasterId, strProServiceId).Copy()
            End If
            tbChooseProduct.TableName = "ChooseProduct"
            dsReturn.Tables.Add(tbChooseProduct)
            Dim tbCanChooseCharge As DataTable = Nothing
            If filterCustId Then
                tbCanChooseCharge = GetCanChooseCharge(CustId).Copy()
            Else
                tbCanChooseCharge = GetNewCanChooseCharge(SEQNO).Copy()
            End If
            tbCanChooseCharge.TableName = "CanChooseCharge"
            dsReturn.Tables.Add(tbCanChooseCharge)
            Dim tbIsAchBank As New DataTable("IsAchBank")
            tbIsAchBank.Columns.Add("ErrorCode", GetType(Integer))
            tbIsAchBank.Columns.Add("ErrorMessage", GetType(String))
            tbIsAchBank.Columns.Add("ResultBoolean", GetType(Integer))
            Dim rwNew As DataRow = tbIsAchBank.NewRow
            rwNew.Item("ErrorCode") = 0
            rwNew.Item("ErrorMessage") = String.Empty
            rwNew.Item("ResultBoolean") = 0
            If (tbAccount IsNot Nothing) AndAlso (tbAccount.Rows.Count > 0) AndAlso
                (Not DBNull.Value.Equals(tbAccount.Rows(0).Item("BankCode"))) Then
                Dim result As RIAResult = IsACHBank(tbAccount.Rows(0).Item("BankCode"))
                rwNew.Item("ErrorCode") = result.ErrorCode
                rwNew.Item("ErrorMessage") = result.ErrorMessage
                rwNew.Item("ResultBoolean") = 0
                If result.ResultBoolean Then
                    rwNew.Item("ResultBoolean") = 1
                End If
            End If
            tbIsAchBank.Rows.Add(rwNew)
            dsReturn.Tables.Add(tbIsAchBank.Copy)
        Catch ex As Exception
            Throw
        Finally
            'DAO.AutoCloseConn = True
            'DAO.CloseConn()
            'DAO.Dispose()
            'DAO = Nothing
        End Try

        Return dsReturn
    End Function
    Public Function QueryNewAllData(ByVal MasterId As Integer, ByVal SEQNO As Integer) As DataSet
        DAO.AutoCloseConn = False
        Dim dsReturn As New DataSet

        Dim strProServiceId As String = "-1"
        Try
            Using tbAccount As DataTable = DAO.ExecQry(_DAL.QueryAccount(), New Object() {MasterId}).Copy
                Using tbSO137 As DataTable = DAO.ExecQry(_DAL.GetSO137(), New Object() {SEQNO}).Copy

                    If (tbAccount IsNot Nothing) AndAlso (tbAccount.Rows.Count > 0) Then
                        'CustId = Integer.Parse(tbAccount.Rows(0).Item("CustId"))

                        If (DBNull.Value.Equals(tbAccount.Rows(0).Item("SnactionDate"))) AndAlso
                              (Not DBNull.Value.Equals(tbAccount.Rows(0).Item("ProServiceID"))) Then
                            strProServiceId = tbAccount.Rows(0).Item("ProServiceID")
                        End If

                    End If
                        tbAccount.TableName = "Account"
                    tbSO137.TableName = "Declared"
                    dsReturn.Tables.Add(tbAccount.Copy)
                    dsReturn.Tables.Add(tbSO137.Copy)

                    Dim ID As String = String.Empty
                    If Not DBNull.Value.Equals(tbSO137.Rows(0).Item("ID")) Then
                        ID = tbSO137.Rows(0).Item("ID")
                    End If
                End Using
                Using tbSystemPara As DataTable = GetSystemPara.Copy
                    tbSystemPara.TableName = "SystemPara"
                    dsReturn.Tables.Add(tbSystemPara.Copy)
                End Using
                Using tbProposer As DataTable = GetNewProposer(SEQNO).Copy
                    tbProposer.TableName = "Proposer"
                    dsReturn.Tables.Add(tbProposer.Copy)
                End Using

                Using tbCanChooseNonePeriod As DataTable = GetCanChooseNonePeriod(SEQNO).Copy
                    tbCanChooseNonePeriod.TableName = "CanChooseNonePeriod"
                    dsReturn.Tables.Add(tbCanChooseNonePeriod.Copy)
                End Using
                Using tbCanChooseProduct = GetNewCanChooseProduct(SEQNO).Copy
                    tbCanChooseProduct.TableName = "CanChooseProduct"
                    dsReturn.Tables.Add(tbCanChooseProduct.Copy)
                End Using
                Using tbCanChooseBillNo = GetCanChooseBillNo(SEQNO).Copy
                    tbCanChooseBillNo.TableName = "GetCanChooseBillNo"
                    dsReturn.Tables.Add(tbCanChooseBillNo.Copy)
                End Using
                Using tbACHTNo As DataTable = GetACHTNo(Integer.Parse(dsReturn.Tables("SystemPara").Rows(0).Item(0)) = 1).Copy
                    tbACHTNo.TableName = "ACHTNo"
                    dsReturn.Tables.Add(tbACHTNo.Copy)
                End Using
                Using tbChooseProduct As DataTable = QueryNewChooseProduct(MasterId, strProServiceId).Copy
                    tbChooseProduct.TableName = "ChooseProduct"
                    dsReturn.Tables.Add(tbChooseProduct.Copy)
                    Using tbOldChooseProduct As DataTable = tbChooseProduct.Copy()
                        tbOldChooseProduct.TableName = "OldChooseProduct"
                        dsReturn.Tables.Add(tbOldChooseProduct)
                    End Using
                End Using
                Using tbMediaCode As DataTable = GetMediaCode.Copy
                    tbMediaCode.TableName = "MediaCode"
                    dsReturn.Tables.Add(tbMediaCode.Copy)
                End Using
                Using tbAcceptName As DataTable = GetAcceptName.Copy
                    tbAcceptName.TableName = "AcceptName"
                    dsReturn.Tables.Add(tbAcceptName.Copy)
                End Using
                Using tbCMCode As DataTable = GetCMCode.Copy
                    tbCMCode.TableName = "CMCode"
                    dsReturn.Tables.Add(tbCMCode.Copy)
                End Using
                Using tbGetBankCode As DataTable = GetBankCode(Integer.Parse(dsReturn.Tables("SystemPara").Rows(0).Item(0)) = 1).Copy
                    tbGetBankCode.TableName = "BankCode"
                    dsReturn.Tables.Add(tbGetBankCode.Copy)
                End Using
                Using tbCardCode As DataTable = GetCardCode.Copy
                    tbCardCode.TableName = "CardCode"
                    dsReturn.Tables.Add(tbCardCode.Copy)
                End Using
                Using tbPTCode As DataTable = GetPTCode.Copy
                    tbPTCode.TableName = "PTCode"
                    dsReturn.Tables.Add(tbPTCode.Copy)
                End Using
                Using tbPriv As DataTable = GetPriv.Copy
                    tbPriv.TableName = "Priv"
                    dsReturn.Tables.Add(tbPriv)
                End Using
                Using bll As New CableSoft.SO.BLL.Utility.Utility(LoginInfo, DAO)
                    Using dtFieldPriv As DataTable = bll.GetFieldPrivMappingData("SO1100G", IIf(MasterId > 0, EditMode.Edit, EditMode.Append))
                        dtFieldPriv.TableName = "FieldPriv"
                        dsReturn.Tables.Add(dtFieldPriv.Copy)
                        dtFieldPriv.Dispose()
                    End Using

                    bll.Dispose()
                End Using
                Dim AccountNo As String = "X"
                Dim BankCode As Integer = -1
                If tbAccount.Rows.Count > 0 Then
                    AccountNo = tbAccount.Rows(0).Item("AccountID")
                    BankCode = Integer.Parse(tbAccount.Rows(0).Item("BankCode") & "")
                End If
                Using tbCanChooseCharge As DataTable = GetNewCanChooseCharge(SEQNO).Copy
                    tbCanChooseCharge.TableName = "CanChooseCharge"
                    dsReturn.Tables.Add(tbCanChooseCharge.Copy)
                End Using
                Using tbIsAchBank As New DataTable("IsAchBank")
                    tbIsAchBank.Columns.Add("ErrorCode", GetType(Integer))
                    tbIsAchBank.Columns.Add("ErrorMessage", GetType(String))
                    tbIsAchBank.Columns.Add("ResultBoolean", GetType(Integer))
                    Dim rwNew As DataRow = tbIsAchBank.NewRow
                    rwNew.Item("ErrorCode") = 0
                    rwNew.Item("ErrorMessage") = String.Empty
                    rwNew.Item("ResultBoolean") = 0
                    If (tbAccount IsNot Nothing) AndAlso (tbAccount.Rows.Count > 0) AndAlso
                        (Not DBNull.Value.Equals(tbAccount.Rows(0).Item("BankCode"))) Then
                        Dim result As RIAResult = IsACHBank(tbAccount.Rows(0).Item("BankCode"))
                        rwNew.Item("ErrorCode") = result.ErrorCode
                        rwNew.Item("ErrorMessage") = result.ErrorMessage
                        rwNew.Item("ResultBoolean") = 0
                        If result.ResultBoolean Then
                            rwNew.Item("ResultBoolean") = 1
                        End If

                    End If
                    tbIsAchBank.Rows.Add(rwNew)

                    dsReturn.Tables.Add(tbIsAchBank.Copy)
                End Using
            End Using
        Catch ex As Exception
            Throw
        Finally
            DAO.AutoCloseConn = True
            DAO.CloseConn()
            DAO.Dispose()
            DAO = Nothing
        End Try

        Return dsReturn
    End Function
    Public Function QueryAllData(ByVal MasterId As Integer, ByVal CustId As Integer) As DataSet
        DAO.AutoCloseConn = False
        Dim dsReturn As New DataSet
        Try
            Using tbAccount As DataTable = DAO.ExecQry(_DAL.QueryAccount(), New Object() {MasterId}).Copy
                If (tbAccount IsNot Nothing) AndAlso (tbAccount.Rows.Count > 0) Then
                    CustId = Integer.Parse(tbAccount.Rows(0).Item("CustId"))
                End If
                tbAccount.TableName = "Account"
                dsReturn.Tables.Add(tbAccount.Copy)

                Using tbSystemPara As DataTable = GetSystemPara.Copy
                    tbSystemPara.TableName = "SystemPara"
                    dsReturn.Tables.Add(tbSystemPara.Copy)
                End Using
                Using tbProposer As DataTable = GetProposer(CustId).Copy
                    tbProposer.TableName = "Proposer"
                    dsReturn.Tables.Add(tbProposer.Copy)
                End Using
                Using tbCanChooseProduct = GetCanChooseProduct(CustId).Copy
                    tbCanChooseProduct.TableName = "CanChooseProduct"
                    dsReturn.Tables.Add(tbCanChooseProduct.Copy)
                End Using
                Using tbACHTNo As DataTable = GetACHTNo.Copy
                    tbACHTNo.TableName = "ACHTNo"
                    dsReturn.Tables.Add(tbACHTNo.Copy)
                End Using
                Using tbChooseProduct As DataTable = QueryChooseProduct(CustId, MasterId).Copy
                    tbChooseProduct.TableName = "ChooseProduct"
                    dsReturn.Tables.Add(tbChooseProduct.Copy)
                End Using
                Using tbMediaCode As DataTable = GetMediaCode.Copy
                    tbMediaCode.TableName = "MediaCode"
                    dsReturn.Tables.Add(tbMediaCode.Copy)
                End Using
                Using tbAcceptName As DataTable = GetAcceptName.Copy
                    tbAcceptName.TableName = "AcceptName"
                    dsReturn.Tables.Add(tbAcceptName.Copy)
                End Using
                Using tbCMCode As DataTable = GetCMCode.Copy
                    tbCMCode.TableName = "CMCode"
                    dsReturn.Tables.Add(tbCMCode.Copy)
                End Using
                Using tbGetBankCode As DataTable = GetBankCode.Copy
                    tbGetBankCode.TableName = "BankCode"
                    dsReturn.Tables.Add(tbGetBankCode.Copy)
                End Using
                Using tbCardCode As DataTable = GetCardCode.Copy
                    tbCardCode.TableName = "CardCode"
                    dsReturn.Tables.Add(tbCardCode.Copy)
                End Using
                Using tbPTCode As DataTable = GetPTCode.Copy
                    tbPTCode.TableName = "PTCode"
                    dsReturn.Tables.Add(tbPTCode.Copy)
                End Using
                Using tbPriv As DataTable = GetPriv.Copy
                    tbPriv.TableName = "Priv"
                    dsReturn.Tables.Add(tbPriv.Copy)
                End Using
                Using tbCanChooseCharge As DataTable = GetCanChooseCharge(CustId).Copy
                    tbCanChooseCharge.TableName = "CanChooseCharge"
                    dsReturn.Tables.Add(tbCanChooseCharge.Copy)
                End Using
                Using tbIsAchBank As New DataTable("IsAchBank")
                    tbIsAchBank.Columns.Add("ErrorCode", GetType(Integer))
                    tbIsAchBank.Columns.Add("ErrorMessage", GetType(String))
                    tbIsAchBank.Columns.Add("ResultBoolean", GetType(Integer))
                    Dim rwNew As DataRow = tbIsAchBank.NewRow
                    rwNew.Item("ErrorCode") = 0
                    rwNew.Item("ErrorMessage") = String.Empty
                    rwNew.Item("ResultBoolean") = 0
                    If (tbAccount IsNot Nothing) AndAlso (tbAccount.Rows.Count > 0) AndAlso
                        (Not DBNull.Value.Equals(tbAccount.Rows(0).Item("BankCode"))) Then
                        Dim result As RIAResult = IsACHBank(tbAccount.Rows(0).Item("BankCode"))
                        rwNew.Item("ErrorCode") = result.ErrorCode
                        rwNew.Item("ErrorMessage") = result.ErrorMessage
                        rwNew.Item("ResultBoolean") = 0
                        If result.ResultBoolean Then
                            rwNew.Item("ResultBoolean") = 1
                        End If

                    End If
                    tbIsAchBank.Rows.Add(rwNew)
                    dsReturn.Tables.Add(tbIsAchBank.Copy)
                End Using
            End Using
        Catch ex As Exception
            Throw
        Finally
            DAO.AutoCloseConn = True
            DAO.CloseConn()
            DAO.Dispose()
            DAO = Nothing
        End Try
       
        Return dsReturn
    End Function
    ''' <summary>
    ''' 查詢帳號資訊
    ''' </summary>
    ''' <param name="MasterId">帳號唯一值</param>
    ''' <returns>SO106 DataTable</returns>
    ''' <remarks></remarks>
    Public Function QueryAccount(ByVal MasterId As Integer) As DataTable
        Return DAO.ExecQry(_DAL.QueryAccount(), New Object() {MasterId})
    End Function
    ''' <summary>
    ''' 查詢已指定設備資訊
    ''' </summary>
    ''' <param name="CustId">客戶編號</param>
    ''' <param name="MasterId">帳號唯一值</param>
    ''' <returns>SO004 DataTable</returns>
    ''' <remarks></remarks>
    Public Function QueryChooseFaci(ByVal CustId As Integer, ByVal MasterId As Integer) As DataTable
        Return DAO.ExecQry(_DAL.QueryChooseFaci, New Object() {MasterId, CustId})
    End Function
    Public Function QueryNewChooseProduct(ByVal MasterId As Integer, ByVal strProServiceID As String) As DataTable
        Return DAO.ExecQry(_DAL.QueryNewChooseProduct(strProServiceID), New Object() {MasterId})
    End Function
    ''' <summary>
    ''' 查詢已指定產品資訊
    ''' </summary>
    ''' <param name="CustId">客戶編號</param>
    ''' <param name="MasterId">帳號唯一值</param>
    ''' <returns>SO003C DataTable</returns>
    ''' <remarks></remarks>
    Public Function QueryChooseProduct(ByVal CustId As Integer, ByVal MasterId As Integer) As DataTable
        Return DAO.ExecQry(_DAL.QueryChooseProduct, New Object() {MasterId, CustId})
    End Function
    ''' <summary>
    ''' 取得可選付款種類
    ''' </summary>
    ''' <returns>CD032 DataTable</returns>
    ''' <remarks></remarks>
    Public Function GetPTCode() As DataTable
        Return DAO.ExecQry(_DAL.GetPTCode)
    End Function
    ''' <summary>
    ''' 取得可選收費方式
    ''' </summary>
    ''' <returns>CD031 DataTable</returns>
    ''' <remarks></remarks>
    Public Function GetCMCode() As DataTable
        Return DAO.ExecQry(_DAL.GetCMCode)
    End Function
    ''' <summary>
    ''' 取得可選銀行別
    ''' </summary>
    ''' <returns>CD018 DataTable</returns>
    ''' <remarks></remarks>
    Public Function GetBankCode() As DataTable
        Return DAO.ExecQry(_DAL.GetBankCode)
    End Function
    Public Function GetBankCode(ByVal blnStartPost As Boolean) As DataTable
        Return DAO.ExecQry(_DAL.GetBankCode(blnStartPost))
    End Function
    ''' <summary>
    ''' 取得可選信用卡別
    ''' </summary>
    ''' <returns>CD037 DataTable</returns>
    ''' <remarks></remarks>
    Public Function GetCardCode() As DataTable
        Return DAO.ExecQry(_DAL.GetCardCode)
    End Function
    ''' <summary>
    '''查詢可選介紹媒介
    ''' </summary>
    ''' <returns>CD009 DataTable</returns>
    ''' <remarks></remarks>
    Public Function GetMediaCode() As DataTable
        Return DAO.ExecQry(_DAL.GetMediaCode)
    End Function

    'Public Function GetIntroId(ByVal MediaRefNo As Integer, ByVal IntroId As String) As DataTable

    '    If MediaRefNo = 1 AndAlso String.IsNullOrEmpty(IntroId) Then
    '        IntroId = "X"
    '    End If

    '    Select Case MediaRefNo
    '        Case 1
    '            Return DAO.ExecQry(_DAL.GetIntroId(MediaRefNo), New Object() {IntroId})
    '        Case Else
    '            Return DAO.ExecQry(_DAL.GetIntroId(MediaRefNo))
    '    End Select


    'End Function
    'Public Function GetIntroData(ByVal MediaRefNo As Integer, ByVal aWhere As String) As DataTable
    '    Return DAO.ExecQry(_DAL.GetIntroData(MediaRefNo) & aWhere)
    '    'Select Case MediaRefNo
    '    '    Case 1

    '    '    Case Else
    '    '        Return DAO.ExecQry(_DAL.GetIntroId(MediaRefNo))
    '    'End Select
    'End Function

    Public Function GetAcceptName() As DataTable
        Return DAO.ExecQry(_DAL.GetAcceptName, Me.LoginInfo.CompCode)
    End Function
    Public Function GetNewCanChooseCharge(ByVal SEQNO As String) As DataTable
        'Return DAO.ExecQry(_DAL.GetNewCanChooseCharge(SEQNO))
        Return DAO.ExecQry(_DAL.GetNewCanChooseCharge, New Object() {Integer.Parse(SEQNO)})
    End Function
    Public Function GetCanChooseCharge(ByVal CustId As Integer) As DataTable
        Return DAO.ExecQry(_DAL.GetCanChooseCharge, New Object() {CustId})
    End Function
    ''' <summary>
    ''' 取得ACH交易別
    ''' </summary>
    ''' <returns>CD068 DataTable</returns>
    ''' <remarks></remarks>
    Public Function GetACHTNo(ByVal blnStartPost As Boolean) As DataTable
        Return DAO.ExecQry(_DAL.GetACHTNo(blnStartPost))
    End Function
    Public Function GetACHTNo() As DataTable
        Return DAO.ExecQry(_DAL.GetACHTNo)
    End Function
    ''' <summary>
    ''' 取得虛擬帳號
    ''' </summary>
    ''' <param name="CustId">CustId</param>
    ''' <param name="BankCode">BankCode</param>
    ''' <returns>ResultXML(虛擬帳號) </returns>
    ''' <remarks></remarks>
    Public Function GetVirtualAccount(ByVal CustId As Integer, ByVal BankCode As Integer) As RIAResult
        Dim intCount As Int32 = 0
        Dim intACTLENGTH As Int32 = 0
        Dim aResult As String = String.Empty
        Try
            intCount = Int32.Parse(DAO.ExecSclr(_DAL.GetVirtualAccountQry, New Object() {CustId}))
            intACTLENGTH = Int32.Parse(DAO.ExecSclr(_DAL.GetActLength, BankCode)) - 8
            aResult = Right("00000000" & intCount.ToString, 8) &
                Right(New String("0"c, intACTLENGTH) & CustId.ToString, intACTLENGTH)
            Return New RIAResult() With {.ResultBoolean = True, .ErrorCode = 0, .ErrorMessage = String.Empty, .ResultXML = aResult}
        Catch ex As Exception
            Return New RIAResult() With {.ErrorCode = -1, .ErrorMessage = ex.Message, .ResultBoolean = False}
        End Try
    End Function
    ''' <summary>
    ''' 取得虛擬帳號
    ''' </summary>
    ''' <param name="CustId">CustId</param>
    ''' <param name="BankCode">BankCode</param>
    ''' <returns>ResultXML(虛擬帳號) </returns>
    ''' <remarks></remarks>
    Public Function GetOldVirtualAccount(ByVal CustId As Integer, ByVal BankCode As Integer) As RIAResult
        Dim intCount As Int32 = 0
        Dim intACTLENGTH As Int32 = 0
        Dim aResult As String = String.Empty
        Try
            intCount = Int32.Parse(DAO.ExecSclr(_DAL.GetOldVirtualAccountQry, New Object() {CustId}))
            intACTLENGTH = Int32.Parse(DAO.ExecSclr(_DAL.GetActLength, BankCode))
            aResult = Right("00000000" & intCount.ToString, 8) & New String("0"c, intACTLENGTH - 8 - CustId.ToString.Length) & CustId.ToString
            'aResult = Right("00000000" & intCount.ToString, 8) &
            '    Right(New String("0"c, intACTLENGTH) & CustId.ToString, intACTLENGTH)
            Return New RIAResult() With {.ResultBoolean = True, .ErrorCode = 0, .ErrorMessage = String.Empty, .ResultXML = aResult}
        Catch ex As Exception
            Return New RIAResult() With {.ErrorCode = -1, .ErrorMessage = ex.Message, .ResultBoolean = False}
        End Try
    End Function
    ''' <summary>
    ''' 取得可指定設備
    ''' </summary>
    ''' <param name="CustId">GetCanChooseFaci</param>
    ''' <returns>SO004 DataTable</returns>
    ''' <remarks></remarks>
    Public Function GetCanChooseFaci(ByVal CustId As Integer) As DataTable
        Return DAO.ExecQry(_DAL.GetCanChooseFaci, New Object() {CustId})
    End Function
    Public Function GetNewCanChooseProdutWithACH(ByVal SEQNO As Integer, ByVal ACHTNO As String, ByVal ACHTDESC As String) As DataSet
        Dim ds As New DataSet
        Dim tb1 As DataTable = Nothing
        Dim tb2 As DataTable = Nothing
        Dim tb3 As DataTable = Nothing
        Try
            tb1 = DAO.ExecQry(_DAL.GetNewCanChooseProdutWithACH(ACHTNO, ACHTDESC), New Object() {SEQNO})
            tb1.TableName = "CHOOSEPRODUCT"
            tb2 = DAO.ExecQry(_DAL.GetNewCanChooseNonePeriodWithACH(ACHTNO, ACHTDESC), New Object() {SEQNO})
            tb2.TableName = "NONEPERIOD"
            tb3 = DAO.ExecQry(_DAL.GetNewCanChooseBillNo(ACHTNO, ACHTDESC), New Object() {SEQNO})
            tb3.TableName = "CHOOSEBILLNO"
            ds.Tables.Add(tb1.Copy)
            ds.Tables.Add(tb2.Copy)
            ds.Tables.Add(tb3.Copy)
            Return ds.Copy
        Catch ex As Exception
            Throw
        Finally
            If tb1 IsNot Nothing Then
                tb1.Dispose()
                tb1 = Nothing
            End If
            If tb2 IsNot Nothing Then
                tb2.Dispose()
                tb2 = Nothing
            End If
            If tb3 IsNot Nothing Then
                tb3.Dispose()
                tb3 = Nothing
            End If
            If ds IsNot Nothing Then
                ds.Dispose()
                ds = Nothing
            End If
        End Try

    End Function
    Public Function GetNewCanChooseProduct(ByVal SeqNo As String) As DataTable
        '        Return DAO.ExecQry(_DAL.GetNewCanChooseProduct(SeqNo))
        Return DAO.ExecQry(_DAL.GetNewCanChooseProduct, New Object() {Integer.Parse(SeqNo)})
    End Function
    Public Function GetCanChooseProduct(ByVal CustId As Integer) As DataTable
        Return DAO.ExecQry(_DAL.GetCanChooseProduct, New Object() {CustId})
    End Function
    ''' <summary>
    ''' 可新增
    ''' </summary>
    ''' <returns>0: 成功 1: 失敗</returns>
    ''' <remarks></remarks>
    Public Function CanAppend() As RIAResult
        Dim obj As New CableSoft.SO.BLL.Utility.Utility(Me.LoginInfo, DAO)
        Try
            Return obj.ChkPriv(Me.LoginInfo.EntryId, "SO1100G1")
        Finally
            obj.Dispose()
        End Try
        'Return New CableSoft.SO.BLL.Utility.Utility(Me.LoginInfo).ChkPriv(Me.LoginInfo.EntryId, "SO1100G1")
    End Function
    Public Function CanView() As CableSoft.BLL.Utility.RIAResult
        Dim obj As New CableSoft.SO.BLL.Utility.Utility(Me.LoginInfo, DAO)
        Try
            Return obj.ChkPriv(Me.LoginInfo.EntryId, "SO1100G4")
        Finally
            obj.Dispose()
        End Try
        'Return New CableSoft.SO.BLL.Utility.Utility(Me.LoginInfo).ChkPriv(Me.LoginInfo.EntryId, "SO1100G4")
    End Function
    ''' <summary>
    ''' 可修改
    ''' </summary>
    ''' <returns>0: 成功 1: 失敗</returns>
    ''' <remarks></remarks>
    Public Function CanEdit() As RIAResult
        Dim obj As New CableSoft.SO.BLL.Utility.Utility(Me.LoginInfo, DAO)
        Try
            Return obj.ChkPriv(Me.LoginInfo.EntryId, "SO1100G2")
        Finally
            obj.Dispose()
        End Try
        'Return New CableSoft.SO.BLL.Utility.Utility(Me.LoginInfo).ChkPriv(Me.LoginInfo.EntryId, "SO1100G2")
    End Function
    ''' <summary>
    ''' 可刪除
    ''' </summary>
    ''' <returns>0: 成功 1: 失敗</returns>
    ''' <remarks></remarks>
    Public Function CanDelete() As RIAResult
        Dim obj As New CableSoft.SO.BLL.Utility.Utility(Me.LoginInfo, DAO)
        Try
            Return obj.ChkPriv(Me.LoginInfo.EntryId, "SO1100G3")
        Finally
            obj.Dispose()
        End Try
        'Return New CableSoft.SO.BLL.Utility.Utility(Me.LoginInfo).ChkPriv(Me.LoginInfo.EntryId, "SO1100G3")
    End Function
    ''' <summary>
    ''' 可列印
    ''' </summary>
    ''' <returns>0: 成功 1: 失敗</returns>
    ''' <remarks></remarks>
    Public Function CanPrint() As RIAResult
        Dim obj As New CableSoft.SO.BLL.Utility.Utility(Me.LoginInfo, DAO)
        Try
            Return obj.ChkPriv(Me.LoginInfo.EntryId, "SO1100G5")
        Finally
            obj.Dispose()
        End Try
        'Return New CableSoft.SO.BLL.Utility.Utility(Me.LoginInfo).ChkPriv(Me.LoginInfo.EntryId, "SO1100G5")
    End Function
    ''' <summary>
    ''' 取得所有權限
    ''' </summary>
    ''' <returns>0: 成功 1: 失敗</returns>
    ''' <remarks></remarks>
    Public Function GetPriv() As DataTable
        Dim obj As New CableSoft.SO.BLL.Utility.Utility(Me.LoginInfo, DAO)
        Try
            Dim dt As DataTable = obj.GetPriv(Me.LoginInfo.EntryId, "SO1100G")
            Return dt
        Finally
            obj.Dispose()
        End Try

        'Return New CableSoft.SO.BLL.Utility.Utility(Me.LoginInfo).ChkPriv(Me.LoginInfo.EntryId, "SO1100G")
    End Function
    Public Function DataOK() As RIAResult
        Return New RIAResult()
    End Function
    Public Function ChkCreditCard(ByVal Account As DataSet) As RIAResult
        Dim result As New RIAResult With {.ErrorMessage = Nothing, .ErrorCode = -704, .ResultBoolean = False}
        If DBNull.Value.Equals(Account.Tables(FNewAccountTableName).Rows(0).Item("CardCode")) Then
            result.ResultBoolean = True
            result.ErrorCode = 0
            Return result
        End If
        With Account.Tables(FNewAccountTableName).Rows(0)
            Using tbCD037 As DataTable = DAO.ExecQry(_DAL.GetCardCodeByCode, New Object() {Integer.Parse(.Item("CardCode"))})
                If tbCD037.Rows.Count = 0 Then
                    result.ErrorMessage = Language.noFoundCardCode
                    Return result
                End If
                If DBNull.Value.Equals(tbCD037.Rows(0).Item("RefNo")) Then
                    result.ErrorMessage = Language.noFoundCardType
                    Return result
                End If
                Select Case Integer.Parse(tbCD037.Rows(0).Item("RefNo"))
                    Case 1
                        If .Item("AccountID").ToString.Substring(0, 1) <> "4" Then result.ErrorMessage = Language.VisaHeader : result.ErrorCode = -1 : Return result
                        If .Item("AccountID").ToString.Length <> 16 Then result.ErrorMessage = Language.VisaLenLimit : Return result
                    Case 2
                        If .Item("AccountID").ToString.Substring(0, 1) <> "5" Then result.ErrorMessage = Language.MasterHeader : result.ErrorCode = -1 : Return result
                        If .Item("AccountID").ToString.Length <> 16 Then result.ErrorMessage = Language.MasterLenLimit : Return result
                    Case 3
                        If .Item("AccountID").ToString.Substring(0, 1) <> "3" Then result.ErrorMessage = Language.JCBHeader : result.ErrorCode = -1 : Return result
                        If .Item("AccountID").ToString.Length <> 16 Then result.ErrorMessage = Language.JCBLenLimit : Return result
                    Case 4
                    Case 5
                        If .Item("AccountID").ToString.Length <> 15 Then result.ErrorMessage = Language.AmericaLimit : Return result
                    Case 5
                        If .Item("AccountID").ToString.Length <> 14 Then result.ErrorMessage = Language.BigLimit : Return result
                End Select
            End Using
        End With
        result.ResultBoolean = True
        result.ErrorCode = 0
        Return result
    End Function
    ''' <summary>
    ''' 檢核資料正確性
    ''' </summary>
    ''' <param name="EditMode">存檔種類</param>
    ''' <param name="Account">SO106</param>
    ''' <returns>0: 成功,-1: 失敗</returns>
    ''' <remarks></remarks>
    Public Function ChkDataOk(ByVal EditMode As EditMode, ByVal Account As DataSet) As RIAResult
        Dim result As New RIAResult With {.ErrorMessage = Nothing, .ErrorCode = 0, .ResultBoolean = True}
        Try
            If Not HaveMustField(Account.Tables(FNewAccountTableName)) Then
                Return New RIAResult() With {.ErrorCode = -1, .ErrorMessage = Language.MustField, .ResultBoolean = False}
            Else
                result = ChkMustCMData(Account.Tables(FNewAccountTableName).Rows(0))
                If result.ResultBoolean Then
                    result = ChkCreditCard(Account)
                    If result.ResultBoolean Then
                        Dim blnStartPos As Boolean = False
                        blnStartPos = Integer.Parse(DAO.ExecSclr(_DAL.IsStartPos)) = 1
                        If (Integer.Parse(DAO.ExecSclr(_DAL.IsACHBank(blnStartPos), New Object() {Account.Tables(FNewAccountTableName).Rows(0).Item("BankCode")})) > 0) AndAlso
                                     (String.IsNullOrEmpty(Account.Tables(FNewAccountTableName).Rows(0).Item("ACHTNo").ToString)) AndAlso
                                      (Integer.Parse("0" & Account.Tables(FNewAccountTableName).Rows(0).Item("StopFlag")) = 0) Then
                            Return New RIAResult() With {.ErrorCode = -1, .ErrorMessage = Language.MustACHT, .ResultBoolean = False}
                        End If

                        If ChkAchSN(EditMode, Account) Then
                            Return New RIAResult() With {.ErrorCode = 0, .ErrorMessage = String.Empty, .ResultBoolean = True}
                        Else
                            Return New RIAResult() With {.ErrorCode = -1, .ErrorMessage = Language.AchSNDouble, .ResultBoolean = False}
                        End If
                    End If
                    
                Else
                    Return result
                End If
                Return result


            End If

        Catch ex As Exception
            Return New RIAResult() With {.ErrorCode = -1, .ErrorMessage = ex.Message, .ResultBoolean = False}
        End Try
    End Function
    Private Function ChkAchSN(ByVal EditMode As EditMode, ByVal Account As DataSet) As Boolean

        Dim aCnt As Integer = 0

        Dim aMasterId As Integer = -1


        Try
            If (Account.Tables(FNewAccountTableName).Rows(0).IsNull("ACHSN")) OrElse
                (String.IsNullOrEmpty(Account.Tables(FNewAccountTableName).Rows(0).Item("ACHSN"))) Then
                Return True
            End If
            If (EditMode = CableSoft.BLL.Utility.EditMode.Edit) Then
                aMasterId = Int32.Parse(Account.Tables(FNewAccountTableName).Rows(0).Item("MasterId"))

            End If
            aCnt = DAO.ExecSclr(_DAL.ChkAchSN, New Object() {Account.Tables(FNewAccountTableName).Rows(0).Item("ACHSN").ToString,
                                                     aMasterId})
            If aCnt > 0 Then
                Return False
            End If
            Return True
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' 作廢
    ''' </summary>
    ''' <param name="Account">Account: 帳號資訊、ChooseFaci: 指定的設備</param>
    ''' <returns>True or False</returns>
    ''' <remarks></remarks>
    Public Function VoidData(ByVal Account As DataSet) As RIAResult
        Dim cn As DbConnection = DAO.GetConn()
        Dim trans As DbTransaction = Nothing
        Dim CSLog As CableSoft.SO.BLL.DataLog.DataLog = Nothing

        Try
            If Not Account.Tables.Contains(FNewAccountTableName) Then
                Return New RIAResult() With {.ErrorCode = -1, .ErrorMessage = Language.MustAccountTable, .ResultBoolean = False}
            End If
            If Not Account.Tables.Contains(FOldProductTableName) Then
                Return New RIAResult() With {.ErrorCode = -1, .ErrorMessage = Language.MustProductTable, .ResultBoolean = False}
            End If
            If Not Account.Tables(FNewAccountTableName).Columns.Contains("StopFlag") Then
                Return New RIAResult() With {.ErrorCode = -1, .ErrorMessage = Language.MustStopField, .ResultBoolean = False}
            Else
                'Account.Tables(FCurrectTableName).Rows(0).Item("StopFlag") = 1
                If (DBNull.Value.Equals(Account.Tables(FNewAccountTableName).Rows(0).Item("StopFlag"))) OrElse
                     (Int32.Parse("0" & Account.Tables(FNewAccountTableName).Rows(0).Item("StopFlag").ToString) <> 1) Then
                    Account.Tables(FNewAccountTableName).Rows(0).Item("StopFlag") = 1
                    'Return New RIAResult() With {.ErrorCode = -1, .ErrorMessage = "資料未被停用", .ResultBoolean = False}
                End If
            End If
            If Not Account.Tables(FNewAccountTableName).Columns.Contains("StopDate") Then
                Return New RIAResult() With {.ErrorCode = -1, .ErrorMessage = Language.MustStopDateField, .ResultBoolean = False}
            Else
                'Account.Tables(FCurrectTableName).Rows(0).Item("StopDate") = Date.Now
                If (DBNull.Value.Equals(Account.Tables(FNewAccountTableName).Rows(0).Item("StopDate"))) Then
                    'Return New RIAResult() With {.ErrorCode = -1, .ErrorMessage = "無停用日期", .ResultBoolean = False}
                    Account.Tables(FNewAccountTableName).Rows(0).Item("StopDate") = Date.Now
                End If
            End If
            If Not HavePK(EditMode.Edit, Account.Tables(FNewAccountTableName)) Then
                Return New RIAResult() With {.ErrorCode = -1, .ErrorMessage = Language.MustMasterId, .ResultBoolean = False}
            End If
            cn.ConnectionString = Me.LoginInfo.ConnectionString
            cn.Open()
            CSLog = New CableSoft.SO.BLL.DataLog.DataLog(Me.LoginInfo, Me.DAO)

            trans = cn.BeginTransaction
            CableSoft.BLL.Utility.Utility.SetClientInfo(Me.DAO, LoginInfo.EntryName)
            Using cmd As System.Data.Common.DbCommand = DAO._factory.CreateCommand()
                cmd.Connection = cn
                cmd.Transaction = trans
                DAO.AutoCloseConn = False
                DAO.Transaction = trans

                ClearSO003(EditMode.Edit, Account, 0, True, cmd)
                'ClearSO004(EditMode.Edit, Account, 0, True, cmd)
                ClearSO003C(EditMode.Edit, Account, 0, True, cmd)
                cmd.Parameters.Clear()
                Dim PKBillNo As String = String.Empty
                For Each rwCharge As DataRow In Account.Tables(VoidBillTableName).Rows
                    If String.IsNullOrEmpty(PKBillNo) Then
                        PKBillNo = rwCharge.Item("BillNo").ToString & rwCharge.Item("Item").ToString
                    Else
                        PKBillNo = PKBillNo & "," & rwCharge.Item("BillNo").ToString & rwCharge.Item("Item").ToString
                    End If
                Next
                'learSO033(Account.Tables(FNewAccountTableName).Rows(0).Item("CitemStr2").ToString)
                ClearSO033(PKBillNo,
                           Account.Tables(FNewAccountTableName).Rows(0).Item("UpdEn"),
                           Account.Tables(FNewAccountTableName).Rows(0).Item("UpdTime"),
                           Account.Tables(FNewAccountTableName).Rows(0).Item("NewUpdTime"))
                cmd.Parameters.Clear()
                'StopSO002A(EditMode.Edit, Account, 0, cmd)
                'cmd.CommandText = String.Format("UPDATE SO106 SET STOPFLAG = 1,STOPDATE = TO_DATE('{0}','yyyymmdd') " & _
                '        " WHERE MASTERID = {1}", Format(Account.Tables(FNewAccountTableName).Rows(0).Item("StopDate"), "yyyyMMdd"),
                'Account.Tables(FNewAccountTableName).Rows(0).Item("MASTERID"))
                'cmd.CommandText = _DAL.VoidSO106Data(Account.Tables(FNewAccountTableName))

                'cmd.ExecuteNonQuery()
                DAO.ExecNqry(_DAL.VoidSO106Data(Account.Tables(FNewAccountTableName)))
                Dim aResult As RIAResult = CSLog.SummaryExpansion(cmd, DataLog.OpType.Update, "SO106", Account.Tables(FNewAccountTableName), Int32.Parse(Integer.Parse(DateTime.Now.ToString("yyyyMMdd"))))
                If Not aResult.ResultBoolean Then
                    Select Case aResult.ErrorCode
                        Case -5
                        Case -6
                            trans.Rollback()
                            Return aResult
                    End Select
                End If
                trans.Commit()
                Return New RIAResult() With {.ErrorCode = 0, .ErrorMessage = String.Empty, .ResultBoolean = True}
            End Using
        Catch ex As Exception
            If trans IsNot Nothing Then
                trans.Rollback()
            End If
            Return New RIAResult() With {.ErrorCode = -1, .ErrorMessage = ex.Message, .ResultBoolean = False}
        Finally
            If trans IsNot Nothing Then
                trans.Dispose()
            End If
            If cn IsNot Nothing Then
                cn.Close()
                cn.Dispose()
            End If
            If CSLog IsNot Nothing Then
                CSLog.Dispose()
            End If
            DAO.AutoCloseConn = True
        End Try
        Return New RIAResult() With {.ErrorCode = 0, .ErrorMessage = String.Empty, .ResultBoolean = True}
    End Function
    Private Function ChkSO002A(ByVal EditMode As EditMode, ByVal Account As DataSet, ByRef cmd As DbCommand) As Boolean
        Dim aTB001 As DataTable = Nothing
        Dim aTB002 As DataTable = Nothing
        Try
            If newFlow Then Return True
            'Dim aRow As DataRow = Account.Tables(FCurrectTableName).Rows(aRowIndex)
            Dim aMasterId As Int32 = -999
            Dim intPayId As Int32 = GetPayId(Account.Tables(FNewAccountTableName).Rows(0))
            Dim ADDCITEMACCOUNT As Integer = 0
            If Not DBNull.Value.Equals(Account.Tables(FNewAccountTableName).Rows(0).Item("ADDCITEMACCOUNT")) Then
                ADDCITEMACCOUNT = Integer.Parse(Account.Tables(FNewAccountTableName).Rows(0).Item("ADDCITEMACCOUNT"))
            End If
            ' aTB001 = DAO.ExecQry(_DAL.GetSO137CustId, New Object() {Account.Tables(FDeclaredTableName).Rows(0).Item("SEQNO")})
            aTB001 = Account.Tables(FNewAccountTableName).Copy
            'cmd.Parameters.Clear()
            'aTB002 = DAO.ExecQry(_DAL.GetNewSO002, New Object() {Me.LoginInfo.CompCode,
            '                                                     Account.Tables(FDeclaredTableName).Rows(0).Item("SEQNO")})
            cmd.Parameters.Clear()
            If aTB001.Rows.Count <= 0 Then
                Throw New Exception(Language.NoSO001Data)
            End If
            'If aTB002.Rows.Count <= 0 Then
            '    Throw New Exception(Language.NoSO002Data)
            'End If
            For Each custRow In aTB001.Rows
                cmd.Parameters.Clear()
                'aTB002 = DAO.ExecQry(_DAL.GetSO002, New Object() {Me.LoginInfo.CompCode, custRow("CustId")})
                'If aTB002.Rows.Count <= 0 Then
                '    Throw New Exception(Language.NoSO002Data)
                'End If
                aTB002 = Account.Tables("INV").Copy
                cmd.Parameters.Clear()
                If (Account.Tables(FNewAccountTableName).Rows(0).IsNull("StopFlag") OrElse
                            Account.Tables(FNewAccountTableName).Rows(0)("StopFlag") = 0) AndAlso
                                (Not Account.Tables(FNewAccountTableName).Rows(0).IsNull("SnactionDate")) Then

                    If DAO.ExecSclr(_DAL.ChkSO002ACnt, New Object() {Account.Tables(FNewAccountTableName).Rows(0)("ACCOUNTID"),
                                                           custRow("CUSTID"),
                                                           Me.LoginInfo.CompCode}) > 0 Then
                        Dim aBankCode As Int32 = -1
                        Dim aWhere As String = " AND 1=1 "

                        If EditMode = CableSoft.BLL.Utility.EditMode.Edit Then
                            If Account.Tables(FNewAccountTableName).Rows(0).HasVersion(DataRowVersion.Original) Then
                                aBankCode = Account.Tables(FNewAccountTableName).Rows(0).Item("BankCode", DataRowVersion.Original)
                            Else
                                aBankCode = Account.Tables(FNewAccountTableName).Rows(0)("BankCode")
                            End If
                            aWhere = " AND BANKCODE = " & aBankCode
                        End If

                        DAO.ExecNqry(_DAL.UpdSO002A & aWhere, New Object() {Account.Tables(FNewAccountTableName).Rows(0)("BankCode"),
                                                                   Account.Tables(FNewAccountTableName).Rows(0)("BankName"), intPayId,
                                                                            Account.Tables(FNewAccountTableName).Rows(0)("AccountId"),
                                                                    Account.Tables(FNewAccountTableName).Rows(0)("CARDNAME"),
                                                                            Account.Tables(FNewAccountTableName).Rows(0)("StopYM"),
                                                                            Account.Tables(FNewAccountTableName).Rows(0)("CVC2"),
                                                                   Account.Tables(FNewAccountTableName).Rows(0)("Note"),
                                                                            Account.Tables(FNewAccountTableName).Rows(0)("CITEMSTR"),
                                                                            Account.Tables(FNewAccountTableName).Rows(0)("CITEMSTR2"), ADDCITEMACCOUNT,
                                                                   custRow("CustId"), Account.Tables(FNewAccountTableName).Rows(0)("AccountId"), Me.LoginInfo.CompCode})
                    Else
                        Dim ChargeAddrNo As Object = DBNull.Value
                        Dim ChargeAddress As Object = DBNull.Value
                        Dim MailAddrNo As Object = DBNull.Value
                        Dim MailAddress As Object = DBNull.Value
                        With Account.Tables("INV").Rows(0)
                            If Account.Tables("INV").Columns.Contains("ChargeAddressNo") Then
                                ChargeAddrNo = Integer.Parse(.Item("ChargeAddressNo"))
                            End If
                            If Account.Tables("INV").Columns.Contains("ChargeAddrNo") Then
                                ChargeAddrNo = Integer.Parse(.Item("ChargeAddrNo"))
                            End If
                            ChargeAddress = DAO.ExecSclr(_DAL.getSO014AddressByAddrNo, New Object() {ChargeAddrNo})
                            If Account.Tables("INV").Columns.Contains("MailAddressNo") Then
                                MailAddrNo = Integer.Parse(.Item("MailAddressNo"))
                            End If
                            If Account.Tables("INV").Columns.Contains("MailAddrNo") Then
                                MailAddrNo = Integer.Parse(.Item("MailAddrNo"))
                            End If
                            MailAddress = DAO.ExecSclr(_DAL.getSO014AddressByAddrNo, New Object() {MailAddrNo})
                        End With
                        DAO.ExecNqry(_DAL.InsSO002A, New Object() {custRow("CustId"),
                                                                   Me.LoginInfo.CompCode, Account.Tables(FNewAccountTableName).Rows(0)("BankCode"),
                                                                   Account.Tables(FNewAccountTableName).Rows(0)("BankName"),
                                                                   intPayId, Account.Tables(FNewAccountTableName).Rows(0)("AccountId"),
                                                                   Account.Tables(FNewAccountTableName).Rows(0)("CardName"),
                                                                   Account.Tables(FNewAccountTableName).Rows(0)("StopYM"),
                                                                 ChargeAddrNo, ChargeAddress,
                                                                   MailAddrNo, MailAddress,
                                                                   Account.Tables(FNewAccountTableName).Rows(0)("CVC2"),
                                                                   Account.Tables(FNewAccountTableName).Rows(0)("Note"),
                                                                   Account.Tables(FNewAccountTableName).Rows(0)("AccountName"),
                                                                   aTB002.Rows(0)("InvNo"), aTB002.Rows(0)("InvTitle"),
                                                                   Nothing, aTB002.Rows(0)("InvoiceType"),
                                                                   Account.Tables(FNewAccountTableName).Rows(0)("CitemStr"),
                                                                   Account.Tables(FNewAccountTableName).Rows(0)("CitemStr2"), ADDCITEMACCOUNT})

                    End If

                End If
                If ADDCITEMACCOUNT = 0 Then
                  
                        StopSO002A(EditMode, Account, 0, True, cmd)

                End If
            Next


            Return True
        Catch ex As Exception
            Throw
        Finally
            If aTB001 IsNot Nothing Then
                aTB001.Dispose()
                aTB001 = Nothing
            End If
            If aTB002 IsNot Nothing Then
                aTB002.Dispose()
                aTB002 = Nothing
            End If
        End Try
    End Function
    Private Function ChkSO002A(ByVal EditMode As EditMode,
                                ByVal aRow As DataRow,
                                ByRef cmd As DbCommand) As Boolean
        Dim aTB001 As DataTable = Nothing
        Dim aTB002 As DataTable = Nothing
        Try

            'Dim aRow As DataRow = Account.Tables(FCurrectTableName).Rows(aRowIndex)
            Dim aMasterId As Int32 = -999
            Dim intPayId As Int32 = GetPayId(aRow)
            aTB001 = DAO.ExecQry(_DAL.GetSO001, New Object() {Me.LoginInfo.CompCode, aRow("CustId")})
            cmd.Parameters.Clear()
            aTB002 = DAO.ExecQry(_DAL.GetSO002, New Object() {Me.LoginInfo.CompCode, aRow("CustId")})
            cmd.Parameters.Clear()
            If aTB001.Rows.Count <= 0 Then
                Throw New Exception(Language.NoSO001Data)
            End If
            If aTB002.Rows.Count <= 0 Then
                Throw New Exception(Language.NoSO002Data)
            End If
            cmd.Parameters.Clear()
            If (aRow.IsNull("StopFlag") OrElse aRow("StopFlag") = 0) AndAlso (Not aRow.IsNull("SnactionDate")) Then

                If DAO.ExecSclr(_DAL.ChkSO002ACnt, New Object() {aRow("ACCOUNTID"),
                                                       aRow("CUSTID"),
                                                       Me.LoginInfo.CompCode}) > 0 Then
                    Dim aBankCode As Int32 = -1
                    Dim aWhere As String = " AND 1=1 "

                    If EditMode = CableSoft.BLL.Utility.EditMode.Edit Then
                        If aRow.HasVersion(DataRowVersion.Original) Then
                            aBankCode = aRow.Item("BankCode", DataRowVersion.Original)
                        Else
                            aBankCode = aRow("BankCode")
                        End If

                        aWhere = " AND BANKCODE = " & aBankCode
                    End If

                    DAO.ExecNqry(_DAL.UpdSO002A & aWhere, New Object() {aRow("BankCode"),
                                                               aRow("BankName"), intPayId, aRow("AccountId"),
                                                                aRow("CARDNAME"), aRow("StopYM"), aRow("CVC2"),
                                                               aRow("Note"), aRow("CITEMSTR"), aRow("CITEMSTR2"),
                                                               aRow("CustId"), aRow("AccountId"), Me.LoginInfo.CompCode})
                Else
                    DAO.ExecNqry(_DAL.InsSO002A, New Object() {aRow("CustId"),
                                                               Me.LoginInfo.CompCode, aRow("BankCode"), aRow("BankName"),
                                                               intPayId, aRow("AccountId"), aRow("CardName"), aRow("StopYM"),
                                                               aTB001.Rows(0)("ChargeAddrNo"), aTB001.Rows(0)("ChargeAddress"),
                                                               aTB001.Rows(0)("MailAddrNo"), aTB001.Rows(0)("MailAddress"),
                                                               aRow("CVC2"), aRow("Note"), aRow("AccountName"),
                                                               aTB002.Rows(0)("InvNo"), aTB002.Rows(0)("InvTitle"),
                                                               aTB002.Rows(0)("InvAddress"), aTB002.Rows(0)("InvoiceType"),
                                                               aRow("CitemStr"), aRow("CitemStr2"), 1})

                End If
            End If

            Return True
        Catch ex As Exception
            Throw
        Finally
            If aTB001 IsNot Nothing Then
                aTB001.Dispose()
            End If
            If aTB002 IsNot Nothing Then
                aTB002.Dispose()
            End If
        End Try
    End Function
    Private Function GetPayId(ByVal aRow As DataRow) As Int32
        Try
            Dim blnVAcc As Boolean
            Dim aCMRef As Int32 = -1
            Using tb As DataTable = DAO.ExecQry(_DAL.GetCMRefNo, New Object() {aRow("CMCode")})
                If tb.Rows.Count > 0 Then
                    If Not DBNull.Value.Equals(tb.Rows(0).Item(0)) Then
                        aCMRef = Integer.Parse(tb.Rows(0).Item(0))
                    End If
                End If
            End Using


            If aCMRef = 2 OrElse aCMRef = 4 OrElse aCMRef = 5 Then
                blnVAcc = False
            Else
                blnVAcc = True
            End If
            If blnVAcc Then
                Return 2
            Else
                If aCMRef = 4 Then
                    If Not aRow.IsNull("CardName") Then
                        Return 1
                    Else
                        Return 0
                    End If
                Else
                    Return 0
                End If
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function
    Private Function StopSO002A(ByVal EditMode As EditMode,
                                ByVal Account As DataSet,
                                ByVal aRowIndex As Integer,
                                ByRef cmd As DbCommand) As Boolean
        Return StopSO002A(EditMode, Account, aRowIndex, True, cmd)
    End Function
    Private Function StopSO002A(ByVal EditMode As EditMode,
                                ByVal Account As DataSet,
                                ByVal aRowIndex As Integer,
                                ByVal filterCustId As Boolean,
                                ByRef cmd As DbCommand) As Boolean
        Try

            Dim aMasterId As Int32 = -999
            Dim aStopDate As Date = Date.Now
            Dim AddCitemAccount As Integer = 0
            cmd.Parameters.Clear()
            If newFlow Then Return True
            If Not DBNull.Value.Equals(Account.Tables(FNewAccountTableName).Rows(aRowIndex).Item("AddCitemAccount")) Then
                AddCitemAccount = Integer.Parse(Account.Tables(FNewAccountTableName).Rows(aRowIndex).Item("AddCitemAccount"))
            End If
            If ((Account.Tables(FNewAccountTableName).Rows(aRowIndex).IsNull("StopFlag") OrElse _
                    Int32.Parse(Account.Tables(FNewAccountTableName).Rows(aRowIndex)("StopFlag"))) = 0) AndAlso _
                    (Not Account.Tables(FNewAccountTableName).Rows(aRowIndex).IsNull("SnactionDate")) Then
                If AddCitemAccount = 1 Then
                    Return True
                End If

            End If

            If EditMode <> CableSoft.BLL.Utility.EditMode.Append Then
                aMasterId = Account.Tables(FNewAccountTableName).Rows(aRowIndex).Item("MasterId")
            End If
            aStopDate = CType(CType(FNowDate.ToString, Date).ToString("yyyy/MM/dd"), Date)
            If Not Account.Tables(FNewAccountTableName).Rows(aRowIndex).IsNull("StopDate") Then
                aStopDate = Date.Parse(Account.Tables(FNewAccountTableName).Rows(aRowIndex).Item("StopDate"))
            End If

            If filterCustId Then
                If DAO.ExecSclr(_DAL.ChkSameAcc, New Object() {Account.Tables(FNewAccountTableName).Rows(aRowIndex).Item("ACCOUNTID"),
                                                          Me.LoginInfo.CompCode, Account.Tables(FNewAccountTableName).Rows(aRowIndex).Item("CUSTID"),
                                                           aMasterId}) > 0 Then
                    Return True
                End If
            Else
                If DAO.ExecSclr(_DAL.ChkNewSameAcc(Account.Tables(FDeclaredTableName).Rows(0).Item("SEQNO")), _
                                                            New Object() {Account.Tables(FNewAccountTableName).Rows(aRowIndex).Item("ACCOUNTID"),
                                                          Me.LoginInfo.CompCode,
                                                           aMasterId}) > 0 Then
                    Return True
                End If
            End If
            cmd.Parameters.Clear()
            If filterCustId Then
                DAO.ExecNqry(_DAL.StopSO002A(filterCustId, ""), New Object() {aStopDate,
                                                      Account.Tables(FNewAccountTableName).Rows(aRowIndex)("ACCOUNTID"), _
                                                       Account.Tables(FNewAccountTableName).Rows(aRowIndex)("CustId"),
                                                       Me.LoginInfo.CompCode})
            Else
                DAO.ExecNqry(_DAL.StopSO002A(filterCustId, Account.Tables(FDeclaredTableName).Rows(0).Item("SEQNO")), New Object() {aStopDate,
                                                     Account.Tables(FNewAccountTableName).Rows(aRowIndex)("ACCOUNTID"), _
                                                      Me.LoginInfo.CompCode})
            End If

            cmd.Parameters.Clear()

            'If (Not aRow.IsNull("InheritKey")) AndAlso (Int32.Parse("0" & aRow("InheritFlag")) = 0) Then
            '    DAO.ExecNqry(_DAL.StopChildSO106, New Object() {aRow("AccountID"),
            '                                                  aRow("INHERITKEY"),
            '                                                    Me.LoginInfo.CompCode})
            'End If


            Return True
        Catch ex As Exception
            Throw
        End Try
    End Function
    Public Function GetCanChooseBillNo(ByVal SeqNo As String) As DataTable
        Return DAO.ExecQry(_DAL.GetCanChooseBillNo(SeqNo))
    End Function
    Public Function GetCanChooseNonePeriod(ByVal SeqNo As String) As DataTable
        Return DAO.ExecQry(_DAL.GetCanChooseNonePeriod(SeqNo))
    End Function

    Public Function GetNewProposer(ByVal SEQNO As Integer) As DataTable
        Return DAO.ExecQry(_DAL.GetNewProposer, New Object() {SEQNO})
    End Function
    Public Function GetSystemPara() As DataTable
        Return DAO.ExecQry(_DAL.GetSystemPara, New Object() {Me.LoginInfo.CompCode})
    End Function
    ''' <summary>
    ''' 取得申請人
    ''' </summary>
    ''' <param name="CustId">客戶編號</param>
    ''' <returns>申請人 Table</returns>
    ''' <remarks></remarks>
    Public Function GetProposer(ByVal CustId As Integer) As DataTable
        Return DAO.ExecQry(_DAL.GetProposer, New Object() {CustId, Me.LoginInfo.CompCode, CustId, Me.LoginInfo.CompCode})
    End Function
    Private Function GetNewCitemStr(ByVal Account As DataSet, ByVal aRowIndex As Int32, ByRef cmd As DbCommand) As String
        Dim aRet As String = String.Empty
        Try
            Dim aSEQNO As String = GetNewSeqNo(Account, cmd)
            'If String.IsNullOrEmpty(aSEQNO) Then
            '    aSEQNO = "'-1X'"
            'End If
            'Dim aSQL As String = String.Format("SELECT SEQNO,CITEMCODE FROM SO003" &
            '    " WHERE CUSTID = {0} AND FACISEQNO IN ( {1} ) AND NVL(STOPFLAG,0)=0",
            '    Account.Tables(FCurrectTableName).Rows(aRowIndex).Item("CustId"),
            '    aSEQNO)
            'cmd.CommandText = aSQL
            'Using dr As DbDataReader = cmd.ExecuteReader
            '    While dr.Read
            '        If String.IsNullOrEmpty(aRet) Then
            '            aRet = "'" & dr.Item("SEQNO") & "'"

            '        Else
            '            aRet = aRet & ",'" & dr.Item("SEQNO") & "'"
            '        End If
            '    End While
            'End Using

            Return aSEQNO

        Catch ex As Exception
            Throw
        End Try
    End Function
    Private Function GetCitemStr(ByVal Account As DataSet, ByRef cmd As DbCommand) As String
        Return GetCitemStr(Account, True, cmd)
    End Function
    Private Function GetOldCitemStr(ByVal Custid As Integer, ByVal CitemStr As String, ByRef cmd As DbCommand) As String
        Dim aSEQNO As String = Nothing

        Dim aSQL As String = String.Format("SELECT SEQNO,CITEMCODE FROM SO003" &
            " WHERE CUSTID = {0} AND CitemCode IN ( {1} ) AND NVL(STOPFLAG,0)=0",
            Custid,
            CitemStr)
        cmd.CommandText = aSQL
        Using dr As DbDataReader = cmd.ExecuteReader
            While dr.Read
                If String.IsNullOrEmpty(aSEQNO) Then
                    aSEQNO = "'" & dr.Item("SEQNO") & "'"

                Else
                    aSEQNO = aSEQNO & ",'" & dr.Item("SEQNO") & "'"
                End If
            End While
        End Using

        Return aSEQNO
    End Function
    Private Function GetCitemStr(ByVal Account As DataSet, ByVal filterCustId As Boolean, ByRef cmd As DbCommand) As String
        Dim aRet As String = String.Empty
        Try
            Dim aSEQNO As String = GetSeqNo(Account, filterCustId, cmd)
            'If String.IsNullOrEmpty(aSEQNO) Then
            '    aSEQNO = "'-1X'"
            'End If
            'Dim aSQL As String = String.Format("SELECT SEQNO,CITEMCODE FROM SO003" &
            '    " WHERE CUSTID = {0} AND FACISEQNO IN ( {1} ) AND NVL(STOPFLAG,0)=0",
            '    Account.Tables(FCurrectTableName).Rows(aRowIndex).Item("CustId"),
            '    aSEQNO)
            'cmd.CommandText = aSQL
            'Using dr As DbDataReader = cmd.ExecuteReader
            '    While dr.Read
            '        If String.IsNullOrEmpty(aRet) Then
            '            aRet = "'" & dr.Item("SEQNO") & "'"

            '        Else
            '            aRet = aRet & ",'" & dr.Item("SEQNO") & "'"
            '        End If
            '    End While
            'End Using

            Return aSEQNO

        Catch ex As Exception
            Throw
        End Try
    End Function
    Public Function GetNewSeqNo(ByVal aDS As DataSet, ByRef cmd As DbCommand) As String
        Dim aServiceIds As String = String.Empty
        Dim aSO137SeqNo As String = "-1"
        Dim aProductCodes As String = "-99"
        Dim aFaciSeqnos As String = "'X'"
        Dim aSeqNos As String = String.Empty
        If aDS.Tables(FDeclaredTableName).Rows.Count > 0 Then
            aSO137SeqNo = aDS.Tables(FDeclaredTableName).Rows(0).Item("SEQNO")
        End If
        Try
            For Each aRw As DataRow In aDS.Tables(FChangeProductTableName).Rows
                If String.IsNullOrEmpty(aServiceIds) Then
                    aServiceIds = aRw.Item("ServiceId")
                Else
                    aServiceIds = aServiceIds & "," & aRw.Item("ServiceId")
                End If
            Next
            If String.IsNullOrEmpty(aServiceIds) Then
                Return String.Empty
            End If
            cmd.Parameters.Clear()
            'cmd.CommandText = String.Format(_DAL.GetNewSO003C, aServiceIds, aSO137SeqNo)
            'Using dr As DbDataReader = cmd.ExecuteReader
            '    While dr.Read
            '        If Not DBNull.Value.Equals(dr("ProductCode")) Then
            '            aProductCodes = aProductCodes & "," & dr("ProductCode")
            '        End If
            '        If Not DBNull.Value.Equals(dr("FaciSeqNo")) Then
            '            aFaciSeqnos = aFaciSeqnos & ",'" & dr("FaciSeqNo") & "'"
            '        End If
            '    End While
            'End Using
            'cmd.Parameters.Clear()
            'cmd.CommandText = String.Format("Select SeqNo,CitemCode From SO003 " & _
            '    " Where CustId = {0} And FaciSeqNo IN ({1}) And NVL(StopFlag,0) = 0 " & _
            '    " And CitemCode In (Select CodeNo From CD019 Where ProductCode IN ({2}))",
            '    aCustId, aFaciSeqnos, aProductCodes)
            'cmd.CommandText = _DAL.GetCitemCode(aCustId, aFaciSeqnos, aProductCodes)

            'Using dr As DbDataReader = cmd.ExecuteReader
            '    While dr.Read
            '        If Not String.IsNullOrEmpty(aSeqNos) Then
            '            aSeqNos = aSeqNos & ",'" & dr.Item("SEQNO") & "'"
            '        Else
            '            aSeqNos = "'" & dr.Item("SEQNO") & "'"
            '        End If
            '    End While
            'End Using


            Return aSeqNos
        Catch ex As Exception
            Throw
        End Try
    End Function
    Public Function GetSeqNo(ByVal aDS As DataSet, ByVal cmd As DbCommand) As String
        Return GetSeqNo(aDS, True, cmd)
    End Function
    Public Function GetSeqNo(ByVal aDS As DataSet, ByVal filterCustId As Boolean, ByRef cmd As DbCommand) As String
        Dim aServiceIds As String = String.Empty
        Dim aCustId As String = "-1"
        Dim aProductCodes As String = "-99"
        Dim aFaciSeqnos As String = "'X'"
        Dim aSeqNos As String = String.Empty
        Dim tbCitemStr As DataTable = Nothing
        Dim tbSO003C As DataTable = Nothing
        If filterCustId Then
            If aDS.Tables(FChangeProductTableName).Rows.Count > 0 Then
                aCustId = aDS.Tables(FChangeProductTableName).Rows(0).Item("CUSTID")
            End If
        End If
        Try

            For Each aRw As DataRow In aDS.Tables(FChangeProductTableName).Rows
                If Not DBNull.Value.Equals(aRw.Item("ServiceId")) Then
                    If String.IsNullOrEmpty(aServiceIds) Then
                        aServiceIds = aRw.Item("ServiceId")
                    Else
                        aServiceIds = aServiceIds & "," & aRw.Item("ServiceId")
                    End If
                End If
              
            Next
            If String.IsNullOrEmpty(aServiceIds) Then
                Return String.Empty
            End If
            If newFlow Then
                tbSO003C = DAO.ExecQry(_DAL.GetNewSO003C(aServiceIds),
                                                     New Object() {aDS.Tables(FDeclaredTableName).Rows(0).Item("SEQNO")})

                If tbSO003C IsNot Nothing Then
                    For Each rw As DataRow In tbSO003C.Rows
                        If Not DBNull.Value.Equals(rw.Item("ProductCode")) Then
                            aProductCodes = aProductCodes & "," & rw.Item("ProductCode")
                        End If
                        If Not DBNull.Value.Equals(rw.Item("FaciSeqNo")) Then
                            aFaciSeqnos = aFaciSeqnos & ",'" & rw.Item("FaciSeqNo") & "'"
                        End If
                    Next
                End If


                If filterCustId Then
                    tbCitemStr = DAO.ExecQry(_DAL.GetCitemCode(aFaciSeqnos, aProductCodes), New Object() {aCustId})

                Else
                    tbCitemStr = DAO.ExecQry(_DAL.GetNewCitemCode(aFaciSeqnos, aProductCodes),
                                             New Object() {aDS.Tables(FDeclaredTableName).Rows(0).Item("SEQNO")})
                End If
            Else
                tbCitemStr = DAO.ExecQry(_DAL.getOldCitemCode(aCustId, aServiceIds))
            End If

            If tbCitemStr IsNot Nothing Then
                For Each rw As DataRow In tbCitemStr.Rows
                    If Not String.IsNullOrEmpty(aSeqNos) Then
                        aSeqNos = aSeqNos & ",'" & rw.Item("SEQNO") & "'"
                    Else
                        aSeqNos = "'" & rw.Item("SEQNO") & "'"
                    End If
                Next
            End If

            'If filterCustId Then
            '    cmd.CommandText = _DAL.GetCitemCode(aCustId, aFaciSeqnos, aProductCodes)
            'Else
            '    cmd.CommandText = _DAL.GetNewCitemCode(aDS.Tables(FDeclaredTableName).Rows(0).Item("SEQNO"), aFaciSeqnos, aProductCodes)
            'End If
            'Using dr As DbDataReader = cmd.ExecuteReader
            '    While dr.Read
            '        If Not String.IsNullOrEmpty(aSeqNos) Then
            '            aSeqNos = aSeqNos & ",'" & dr.Item("SEQNO") & "'"
            '        Else
            '            aSeqNos = "'" & dr.Item("SEQNO") & "'"
            '        End If
            '    End While
            'End Using


            Return aSeqNos
        Catch ex As Exception
            Throw
        Finally
            If tbCitemStr IsNot Nothing Then
                tbCitemStr.Dispose()
                tbCitemStr = Nothing
            End If
            If tbSO003C IsNot Nothing Then
                tbSO003C.Dispose()
                tbSO003C = Nothing
            End If
        End Try
    End Function
    Private Function GetInvSeqNo(ByVal rwAccount As DataRow) As String
        Dim aRet As String = Nothing
        If Not DBNull.Value.Equals(rwAccount.Item("INVSEQNO")) Then
            Return Nothing
        End If

        Using tbSO002 As DataTable = DAO.ExecQry(_DAL.GetInvSeqNo,
                                                 New Object() {rwAccount.Item("CUSTID")})
            If tbSO002.Rows.Count > 0 Then
                aRet = tbSO002.Rows(0).Item("INVSEQNO")
            Else

            End If
        End Using
        Return aRet
    End Function
    Public Function insSO138(Account As DataSet) As Boolean
        If newFlow Then Return True

        SO138_InvSeqNo = Integer.Parse(DAO.ExecSclr(_DAL.getInvSO138Seqno))
        Dim o As Integer = SO138_InvSeqNo
        Dim InvPurposeCode As Object = Nothing
        Dim InvPurposeName As Object = Nothing
        Dim DenRecCode As Object = Nothing
        Dim DenRecName As Object = Nothing
        Dim DenRecDate As Object = Nothing
        Dim ApplyInvDate As Object = Nothing
        Try
            With Account.Tables("INV").Rows(0)
                Dim ChargeAddress As String = DAO.ExecSclr(_DAL.getSO014AddressByAddrNo, New Object() {Integer.Parse(.Item("ChargeAddressNo"))})
                Dim MailAddress As String = DAO.ExecSclr(_DAL.getSO014AddressByAddrNo, New Object() {Integer.Parse(.Item("MailAddressNo"))})
                If Not DBNull.Value.Equals(.Item("InvPurposeCode")) Then
                    InvPurposeCode = Integer.Parse(.Item("InvPurposeCode"))
                    InvPurposeName = DAO.ExecSclr(_DAL.getInvPurposeNameByCode, New Object() {Integer.Parse(.Item("InvPurposeCode"))})
                End If
                If Not DBNull.Value.Equals(.Item("DenRecCode")) Then
                    DenRecCode = Integer.Parse(.Item("DenRecCode"))
                    DenRecName = DAO.ExecSclr(_DAL.getDenRecNameByCode, New Object() {Integer.Parse(.Item("DenRecCode"))})
                End If
                If Not DBNull.Value.Equals(.Item("DenRecDate")) Then
                    DenRecDate = Date.Parse(.Item("DenRecDate"))
                End If
                If Not DBNull.Value.Equals(.Item("ApplyInvDate")) Then
                    ApplyInvDate = Date.Parse(.Item("ApplyInvDate"))
                End If
                DAO.ExecNqry(_DAL.insertSO138, New Object() {
                        SO138_InvSeqNo, .Item("ChargeTitle"), Integer.Parse(.Item("InvoiceType")), .Item("INVNo"), _
                        .Item("InvTitle"), InvPurposeCode, InvPurposeName, Integer.Parse(.Item("PreInvoice")), _
                        Integer.Parse(.Item("BillMailKind")), DenRecCode, DenRecName, DenRecDate, _
                        .Item("LoveNum"), Integer.Parse(.Item("InvoiceKind")), ApplyInvDate, Integer.Parse(.Item("ChargeAddressNo")), ChargeAddress, _
                        Integer.Parse(.Item("MailAddressNo")), MailAddress, CableSoft.BLL.Utility.DateTimeUtility.GetDTString(FNowDate), FNowDate, LoginInfo.EntryName
                    })
            End With

            With Account.Tables(FNewAccountTableName).Rows(0)
                DAO.ExecNqry(_DAL.insertSO002AD, New Object() {
                             .Item("AccountId"), LoginInfo.CompCode, Integer.Parse(.Item("CustId")), SO138_InvSeqNo})
            End With

        Catch ex As Exception
            Throw ex
        End Try


        Return True
    End Function
    Public Function SaveNewData(ByVal EditMode As EditMode, Account As DataSet) As DataSet
        Dim cn As DbConnection = DAO.GetConn()
        Dim trans As DbTransaction = Nothing
        Dim CSLog As CableSoft.SO.BLL.DataLog.DataLog = Nothing
        Dim blnAutoClose As Boolean = False
        Dim tbLogSO106 As DataTable = Nothing
        'Dim CSLog As New CableSoft.SO.BLL.DataLog.DataLog(Me.LoginInfo)
        'Dim LogResult As CableSoft.BLL.Utility.RIAResult = CSLog.Summary(cmd, BLL.DataLog.OpType.Update, TableName, LogT.Tables(0), Integer.Parse(DateTime.Now.ToString("yyyyMMdd")))
        Dim aAccountTB As DataTable = Nothing
        Dim dsResult As DataSet = Nothing
        FNowDate = Date.Now
        Try
            If Not HavePK(EditMode, Account.Tables(FNewAccountTableName)) Then
                Throw New Exception("SO106 NO PKField")
                'Return New RIAResult() With {.ErrorCode = -1, .ErrorMessage = "SO106 NO PKField", .ResultBoolean = False}
            End If
            aAccountTB = GetCorrectAccountTable(EditMode, Account)

            CSLog = New CableSoft.SO.BLL.DataLog.DataLog(Me.LoginInfo, Me.DAO)

            If DAO.Transaction IsNot Nothing Then

                trans = DAO.Transaction
            Else
                If cn IsNot Nothing AndAlso cn.State <> ConnectionState.Open Then
                    cn.ConnectionString = Me.LoginInfo.ConnectionString
                    cn.Open()
                End If
                trans = cn.BeginTransaction
                DAO.Transaction = trans
                blnAutoClose = True
            End If
            DAO.AutoCloseConn = False
            If blnAutoClose Then
                Dim aAction As String = Nothing
                Select Case EditMode
                    Case CableSoft.BLL.Utility.EditMode.Edit
                        aAction = Language.EditClientInfo
                    Case CableSoft.BLL.Utility.EditMode.Append
                        aAction = Language.AddClientInfo
                    Case Else
                        aAction = Language.EditClientInfo
                End Select
                CableSoft.BLL.Utility.Utility.SetClientInfo(Me.DAO, LoginInfo.EntryId, aAction)
            End If

            FNowDate = DateTime.Parse(DAO.ExecSclr(_DAL.GetSysDate))

            Using cmd As System.Data.Common.DbCommand = DAO._factory.CreateCommand()
                cmd.Connection = cn
                cmd.Transaction = trans
                'DAO.Transaction = trans
                Dim aWhere As String = String.Empty
                If EditMode = CableSoft.BLL.Utility.EditMode.Append Then
                    insSO138(Account)
                End If
                For i As Integer = 0 To aAccountTB.Rows.Count - 1
                    If tbLogSO106 IsNot Nothing Then
                        tbLogSO106.Dispose()
                        tbLogSO106 = Nothing
                    End If
                    Dim NonePeriod As String = aAccountTB.Rows(i).Item("CitemStr") & ""
                    Dim aryNonePeriod As New List(Of String)
                    Dim aryCitemStr As New List(Of String)
                    If Not String.IsNullOrEmpty(NonePeriod) Then
                        aryNonePeriod = NonePeriod.Split(",").ToList
                    End If
                    Dim aCitemStr As String = GetCitemStr(Account, IIf(newFlow, False, True), cmd)
                    If Not String.IsNullOrEmpty(aCitemStr) Then
                        aryCitemStr = aCitemStr.Split(",").ToList
                    End If
                    For Each item As String In aryNonePeriod
                        If Not aryCitemStr.Contains(item) Then
                            aryCitemStr.Add(item)
                        End If
                    Next
                    aCitemStr = String.Join(",", aryCitemStr.ToArray())
                    If Not String.IsNullOrEmpty(aCitemStr) Then
                        aAccountTB.Rows(i).Item("CitemStr") = aCitemStr
                    Else
                        aAccountTB.Rows(i).Item("CitemStr") = DBNull.Value
                    End If
                    aAccountTB.Rows(i).Item("UpdTime") = CableSoft.BLL.Utility.DateTimeUtility.GetDTString(FNowDate)
                    aAccountTB.Rows(i).Item("NewUpdTime") = FNowDate                    
                    If DBNull.Value.Equals(aAccountTB.Rows(i).Item("AcceptTime")) Then
                        aAccountTB.Rows(i).Item("AcceptTime") = FNowDate
                    End If
                    If chkStopFlag(aAccountTB.Rows(i)) Then
                        aAccountTB.Rows(i).Item("ACHTNo") = DBNull.Value
                        aAccountTB.Rows(i).Item("ACHTDESC") = DBNull.Value
                    End If
                    cmd.Parameters.Clear()
                    '抓取SO002的InvSeqNo填入，避免提回出錯 By Kin 2012/09/06

                    'Dim aInvSeqNo As String = Nothing
                    'If DBNull.Value.Equals(aAccountTB.Rows(i).Item("INVSEQNO")) Then
                    '    aInvSeqNo = GetInvSeqNo(aAccountTB.Rows(i))
                    'End If
                    'If Not String.IsNullOrEmpty(aInvSeqNo) Then
                    '    aAccountTB.Rows(i).Item("INVSEQNO") = aInvSeqNo
                    'End If
                    If Not newFlow AndAlso EditMode = CableSoft.BLL.Utility.EditMode.Append Then
                        aAccountTB.Rows(i).Item("InvseqNo") = SO138_InvSeqNo
                    Else
                        If Not DBNull.Value.Equals(aAccountTB.Rows(i).Item("InvseqNo")) Then
                            SO138_InvSeqNo = aAccountTB.Rows(i).Item("InvseqNo")
                        Else
                            SO138_InvSeqNo = DBNull.Value
                        End If

                    End If
                    Select Case EditMode
                        Case CableSoft.BLL.Utility.EditMode.Append

                            If Not DAO.GetInsertOrUpdateCommand(CableSoft.Utility.DataAccess.UpdateMode.InsertRow, aAccountTB, "SO106", i, cmd, aWhere) Then
                                Throw New Exception(Language.InsertSO106Error)
                                'Return New RIAResult() With {.ErrorCode = -1, .ErrorMessage = Lang.InsertSO106Error, .ResultBoolean = False}
                            End If
                        Case CableSoft.BLL.Utility.EditMode.Edit, CableSoft.BLL.Utility.EditMode.Delete
                            'tbLogSO106 = DAO.ExecQry("Select A.rowid,A.* From SO106  A where Masterid = " & Account.Tables(FNewAccountTableName).Rows(i).Item("MasterId"))
                            tbLogSO106 = DAO.ExecQry(_DAL.QuerySO106Log, New Object() {Account.Tables(FNewAccountTableName).Rows(i).Item("MasterId")})
                            'If Account.Tables(FNewAccountTableName).Columns.Contains("RowId") Then
                            '    aWhere = String.Format("ROWID='{0}'",
                            '                       Account.Tables(FNewAccountTableName).Rows(i).Item("RowId"))
                            'Else
                            '    aWhere = String.Format("MasterId={0}",
                            '                          Account.Tables(FNewAccountTableName).Rows(i).Item("MasterId"))

                            'End If
                            If Account.Tables(FNewAccountTableName).Columns.Contains("ROWID") Then
                                aWhere = String.Format("ROWID='{0}'",
                                                   Account.Tables(FNewAccountTableName).Rows(i).Item("ROWID"))
                            Else
                                aWhere = String.Format("MasterId={0}",
                                                      Account.Tables(FNewAccountTableName).Rows(i).Item("MasterId"))

                            End If
                            If EditMode = CableSoft.BLL.Utility.EditMode.Edit Then
                                If Not DAO.GetInsertOrUpdateCommand(CableSoft.Utility.DataAccess.UpdateMode.UpdateRow, aAccountTB, "SO106", i, cmd, aWhere) Then
                                    Throw New Exception(Language.UpdateSO106Error)
                                    'Return New RIAResult() With {.ErrorCode = -1, .ErrorMessage = Lang.UpdateSO106Error, .ResultBoolean = False}
                                End If
                            Else
                                ' cmd.CommandText = "DELETE SO106 WHERE " & aWhere
                                DAO.ExecNqry(_DAL.DeleteSO106(aWhere))
                                'cmd.CommandText = _DAL.DeleteSO106(aWhere)
                            End If

                    End Select
                    Dim aVoidData As Boolean = False


                    aVoidData = chkStopFlag(aAccountTB.Rows(i))
                    cmd.ExecuteNonQuery()
                    cmd.Parameters.Clear()
                    UpdSO003C(EditMode, Account, i, cmd)
                    cmd.Parameters.Clear()
                    If newFlow Then
                        If Not IsACHBank(Account.Tables(FNewAccountTableName).Rows(0).Item("BankCode").ToString).ResultBoolean Then
                            ClearSO003C(EditMode, Account, i, aVoidData, cmd)
                        Else
                            If chkSnactionDate(Account.Tables(FNewAccountTableName).Rows(0)) Then
                                ClearSO003C(EditMode, Account, i, aVoidData, cmd)
                            End If
                        End If
                    End If


                    'UpdSO004(EditMode, Account, i, cmd)
                    'ClearSO004(EditMode, Account, i, False, cmd)
                    cmd.Parameters.Clear()
                    UpdSO003(EditMode, Account, i, IIf(newFlow, False, True), cmd)
                    Dim oldCitemStr As String = ""
                    If tbLogSO106 IsNot Nothing Then
                        If Not DBNull.Value.Equals(tbLogSO106.Rows(0).Item("CitemStr")) Then
                            oldCitemStr = tbLogSO106.Rows(0).Item("CitemStr")
                        End If

                    End If
                    ClearNonePeriod(EditMode, Account, tbLogSO106, oldCitemStr, cmd)
                    UpdNonePeriod(EditMode, Account, i, Account.Tables(FNewAccountTableName).Copy, NonePeriod, cmd)
                    cmd.Parameters.Clear()
                    ClearSO003(EditMode, Account, i, aVoidData, IIf(newFlow, False, True), cmd)
                    'ClearOldSO003(EditMode, Account, Account.Tables(FNewAccountTableName), _
                    '              Account.Tables(FNewAccountTableName).Rows(0).Item("CitemStr"), _
                    '               cmd)
                    cmd.Parameters.Clear()
                    StopSO002A(EditMode, Account, i, True, cmd)
                    cmd.Parameters.Clear()
                    ChkSO002A(EditMode, Account, cmd)
                    cmd.Parameters.Clear()
                    UpdSO033(EditMode, tbLogSO106, Account, i, cmd)
                    cmd.Parameters.Clear()
                    ProcessACH(EditMode, Account, i, False, cmd)
                    cmd.Parameters.Clear()
                    Dim PKBillNo As String = String.Empty
                    For Each rwCharge As DataRow In Account.Tables(VoidBillTableName).Rows
                        If String.IsNullOrEmpty(PKBillNo) Then
                            PKBillNo = rwCharge.Item("BillNo").ToString & rwCharge.Item("Item").ToString
                        Else
                            PKBillNo = PKBillNo & "," & rwCharge.Item("BillNo").ToString & rwCharge.Item("Item").ToString
                        End If
                    Next
                    'learSO033(Account.Tables(FNewAccountTableName).Rows(0).Item("CitemStr2").ToString)
                    ClearSO033(PKBillNo, aAccountTB.Rows(i).Item("UpdEn"), aAccountTB.Rows(i).Item("UpdTime"), aAccountTB.Rows(i).Item("NewUpdTime"))
                    stopOldData(Account)
                    cmd.Parameters.Clear()
                    'InsSO106Log(EditMode, Account, i, cmd)
                    cmd.Parameters.Clear()
                    If tbLogSO106 IsNot Nothing OrElse 1 = 0 Then
                        CableSoft.BLL.Utility.Utility.CopyDataRow(Account.Tables(FNewAccountTableName).Rows(0), tbLogSO106.Rows(0))
                        'Dim aResult As RIAResult = CSLog.SummaryExpansion(cmd, DataLog.OpType.Update, "SO106", Account.Tables(FNewAccountTableName), Int32.Parse(Integer.Parse(DateTime.Now.ToString("yyyyMMdd"))))
                        Dim aResult As RIAResult = CSLog.SummaryExpansion(cmd, DataLog.OpType.Update, "SO106", tbLogSO106, Int32.Parse(Integer.Parse(DateTime.Now.ToString("yyyyMMdd"))))
                        If Not aResult.ResultBoolean Then
                            Select Case aResult.ErrorCode
                                Case -5
                                Case -6
                                    If blnAutoClose Then
                                        trans.Rollback()
                                        Throw New Exception(aResult.ErrorMessage)
                                        'Return aResult
                                    End If

                            End Select

                        End If
                    End If

                Next
                If newFlow Then
                    dsResult = QueryReLoadData(Integer.Parse(Account.Tables(FNewAccountTableName).Rows(0).Item("MasterId")),
                                                    Account.Tables(FDeclaredTableName).Rows(0).Item("SEQNO"),
                                                    False)
                Else
                    dsResult = QueryOldReLoadData(Integer.Parse(Account.Tables(FNewAccountTableName).Rows(0).Item("MasterId")))
                End If

                If blnAutoClose Then
                    trans.Commit()
                End If

            End Using
            Return dsResult
        Catch ex As Exception
            If (trans IsNot Nothing) AndAlso (blnAutoClose) Then
                trans.Rollback()
            End If
            Throw
            'Return New RIAResult() With {.ErrorCode = -1, .ErrorMessage = ex.Message, .ResultBoolean = False}
        Finally
            If blnAutoClose Then
                CableSoft.BLL.Utility.Utility.ClearClientInfo(DAO)
                DAO.AutoCloseConn = True
                If trans IsNot Nothing Then
                    trans.Dispose()
                End If
                If cn IsNot Nothing Then
                    cn.Close()
                    cn.Dispose()
                End If

            End If
            If tbLogSO106 IsNot Nothing Then
                tbLogSO106.Dispose()
                tbLogSO106 = Nothing
            End If
            If CSLog IsNot Nothing Then
                CSLog.Dispose()
                CSLog = Nothing
            End If
            If aAccountTB IsNot Nothing Then
                aAccountTB.Dispose()
                aAccountTB = Nothing
            End If
        End Try
    End Function

    Public Function SaveData(ByVal EditMode As EditMode, Account As DataSet) As DataSet
        Dim cn As DbConnection = DAO.GetConn()
        Dim trans As DbTransaction = Nothing
        Dim CSLog As CableSoft.SO.BLL.DataLog.DataLog = Nothing
        Dim blnAutoClose As Boolean = False
        'Dim CSLog As New CableSoft.SO.BLL.DataLog.DataLog(Me.LoginInfo)
        'Dim LogResult As CableSoft.BLL.Utility.RIAResult = CSLog.Summary(cmd, BLL.DataLog.OpType.Update, TableName, LogT.Tables(0), Integer.Parse(DateTime.Now.ToString("yyyyMMdd")))
        Dim aAccountTB As DataTable = Nothing
        Dim dsResult As DataSet = Nothing
        FNowDate = Date.Now
        Try
            If Not HavePK(EditMode, Account.Tables(FNewAccountTableName)) Then
                Throw New Exception("SO106 NO PKField")
                'Return New RIAResult() With {.ErrorCode = -1, .ErrorMessage = "SO106 NO PKField", .ResultBoolean = False}
            End If
            aAccountTB = GetCorrectAccountTable(EditMode, Account)

            CSLog = New CableSoft.SO.BLL.DataLog.DataLog(Me.LoginInfo, Me.DAO)

            If DAO.Transaction IsNot Nothing Then

                trans = DAO.Transaction
            Else
                If cn IsNot Nothing AndAlso cn.State <> ConnectionState.Open Then
                    cn.ConnectionString = Me.LoginInfo.ConnectionString
                    cn.Open()
                End If
                trans = cn.BeginTransaction
                DAO.Transaction = trans
                blnAutoClose = True
            End If
            DAO.AutoCloseConn = False
            CableSoft.BLL.Utility.Utility.SetClientInfo(Me.DAO, LoginInfo.EntryName)
            FNowDate = DateTime.Parse(DAO.ExecSclr(_DAL.GetSysDate))

            Using cmd As System.Data.Common.DbCommand = DAO._factory.CreateCommand()
                cmd.Connection = cn
                cmd.Transaction = trans
                'DAO.Transaction = trans
                Dim aWhere As String = String.Empty

                For i As Integer = 0 To aAccountTB.Rows.Count - 1
                    Dim aCitemStr As String = GetCitemStr(Account, cmd)
                    If Not String.IsNullOrEmpty(aCitemStr) Then
                        aAccountTB.Rows(i).Item("CitemStr") = aCitemStr
                    Else
                        aAccountTB.Rows(i).Item("CitemStr") = DBNull.Value
                    End If
                    aAccountTB.Rows(i).Item("UpdTime") = CableSoft.BLL.Utility.DateTimeUtility.GetDTString(FNowDate)
                    aAccountTB.Rows(i).Item("NewUpdTime") = FNowDate
                    If chkStopFlag(aAccountTB.Rows(i)) Then
                        aAccountTB.Rows(i).Item("ACHTNo") = DBNull.Value
                        aAccountTB.Rows(i).Item("ACHTDESC") = DBNull.Value
                    End If
                    cmd.Parameters.Clear()
                    '抓取SO002的InvSeqNo填入，避免提回出錯 By Kin 2012/09/06
                    Dim aInvSeqNo As String = Nothing
                    If DBNull.Value.Equals(aAccountTB.Rows(i).Item("INVSEQNO")) Then
                        aInvSeqNo = GetInvSeqNo(aAccountTB.Rows(i))
                    End If
                    If Not String.IsNullOrEmpty(aInvSeqNo) Then
                        aAccountTB.Rows(i).Item("INVSEQNO") = aInvSeqNo
                    End If
                    Select Case EditMode
                        Case CableSoft.BLL.Utility.EditMode.Append
                            If Not DAO.GetInsertOrUpdateCommand(CableSoft.Utility.DataAccess.UpdateMode.InsertRow, aAccountTB, "SO106", i, cmd, aWhere) Then
                                Throw New Exception(Language.InsertSO106Error)
                                'Return New RIAResult() With {.ErrorCode = -1, .ErrorMessage = Lang.InsertSO106Error, .ResultBoolean = False}
                            End If
                        Case CableSoft.BLL.Utility.EditMode.Edit, CableSoft.BLL.Utility.EditMode.Delete
                            'If Account.Tables(FNewAccountTableName).Columns.Contains("RowId") Then
                            '    aWhere = String.Format("ROWID='{0}'",
                            '                       Account.Tables(FNewAccountTableName).Rows(i).Item("RowId"))
                            'Else
                            '    aWhere = String.Format("MasterId={0}",
                            '                          Account.Tables(FNewAccountTableName).Rows(i).Item("MasterId"))

                            'End If
                            If Account.Tables(FNewAccountTableName).Columns.Contains("Rowid") Then
                                aWhere = String.Format("Rowid='{0}'",
                                                   Account.Tables(FNewAccountTableName).Rows(i).Item("Rowid"))
                            Else
                                aWhere = String.Format("MasterId={0}",
                                                      Account.Tables(FNewAccountTableName).Rows(i).Item("MasterId"))

                            End If
                            If EditMode = CableSoft.BLL.Utility.EditMode.Edit Then
                                If Not DAO.GetInsertOrUpdateCommand(CableSoft.Utility.DataAccess.UpdateMode.UpdateRow, aAccountTB, "SO106", i, cmd, aWhere) Then
                                    Throw New Exception(Language.UpdateSO106Error)
                                    'Return New RIAResult() With {.ErrorCode = -1, .ErrorMessage = Lang.UpdateSO106Error, .ResultBoolean = False}
                                End If
                            Else
                                ' cmd.CommandText = "DELETE SO106 WHERE " & aWhere
                                DAO.ExecNqry(_DAL.DeleteSO106(aWhere))
                                'cmd.CommandText = _DAL.DeleteSO106(aWhere)
                            End If

                    End Select
                    Dim aVoidData As Boolean = False

                    aVoidData = chkStopFlag(aAccountTB.Rows(i))
                    cmd.ExecuteNonQuery()
                    cmd.Parameters.Clear()
                    UpdSO003C(EditMode, Account, i, cmd)
                    cmd.Parameters.Clear()
                    ClearSO003C(EditMode, Account, i, aVoidData, cmd)
                    'UpdSO004(EditMode, Account, i, cmd)
                    'ClearSO004(EditMode, Account, i, False, cmd)
                    cmd.Parameters.Clear()
                    UpdSO003(EditMode, Account, i, cmd)
                    cmd.Parameters.Clear()
                    ClearSO003(EditMode, Account, i, aVoidData, cmd)
                    cmd.Parameters.Clear()
                    StopSO002A(EditMode, Account, i, cmd)
                    cmd.Parameters.Clear()
                    ChkSO002A(EditMode, Account.Tables(FNewAccountTableName).Rows(i), cmd)
                    cmd.Parameters.Clear()
                    'UpdSO033(EditMode, Account, i, cmd)
                    cmd.Parameters.Clear()
                    ProcessACH(EditMode, Account, i, cmd)
                    cmd.Parameters.Clear()
                    Dim PKBillNo As String = String.Empty
                    For Each rwCharge As DataRow In Account.Tables(VoidBillTableName).Rows
                        If String.IsNullOrEmpty(PKBillNo) Then
                            PKBillNo = rwCharge.Item("BillNo").ToString & rwCharge.Item("Item").ToString
                        Else
                            PKBillNo = PKBillNo & "," & rwCharge.Item("BillNo").ToString & rwCharge.Item("Item").ToString
                        End If
                    Next
                    'learSO033(Account.Tables(FNewAccountTableName).Rows(0).Item("CitemStr2").ToString)
                    ClearSO033(PKBillNo, Account.Tables(FNewAccountTableName).Rows(0).Item("UpdEn"),
                               Account.Tables(FNewAccountTableName).Rows(0).Item("UpdTime"),
                               Account.Tables(FNewAccountTableName).Rows(0).Item("NewUpdTime"))
                    cmd.Parameters.Clear()
                    InsSO106Log(EditMode, Account, i, cmd)

                    cmd.Parameters.Clear()
                    Dim aResult As RIAResult = CSLog.SummaryExpansion(cmd, DataLog.OpType.Update, "SO106", Account.Tables(FNewAccountTableName), Int32.Parse(Integer.Parse(DateTime.Now.ToString("yyyyMMdd"))))
                    If Not aResult.ResultBoolean Then
                        Select Case aResult.ErrorCode
                            Case -5
                            Case -6
                                If blnAutoClose Then
                                    trans.Rollback()
                                    Throw New Exception(aResult.ErrorMessage)
                                    'Return aResult
                                End If

                        End Select

                    End If
                Next

                dsResult = QueryReLoadData(Integer.Parse(Account.Tables(FNewAccountTableName).Rows(0).Item("MasterId")))
                If blnAutoClose Then
                    trans.Commit()
                End If

            End Using
            Return dsResult
        Catch ex As Exception
            If (trans IsNot Nothing) AndAlso (blnAutoClose) Then
                trans.Rollback()
            End If
            Throw
            'Return New RIAResult() With {.ErrorCode = -1, .ErrorMessage = ex.Message, .ResultBoolean = False}
        Finally
            If blnAutoClose Then
                DAO.AutoCloseConn = True
                If trans IsNot Nothing Then
                    trans.Dispose()
                End If
                If cn IsNot Nothing Then
                    cn.Close()
                    cn.Dispose()
                End If

            End If
            If CSLog IsNot Nothing Then
                CSLog.Dispose()
                CSLog = Nothing
            End If
            If aAccountTB IsNot Nothing Then
                aAccountTB.Dispose()
                aAccountTB = Nothing
            End If
        End Try
    End Function


    Public Function Save(ByVal EditMode As EditMode, Account As DataSet) As RIAResult
        Dim cn As DbConnection = DAO.GetConn()
        Dim trans As DbTransaction = Nothing
        Dim CSLog As CableSoft.SO.BLL.DataLog.DataLog = Nothing
        Dim blnAutoClose As Boolean = False
        'Dim CSLog As New CableSoft.SO.BLL.DataLog.DataLog(Me.LoginInfo)
        'Dim LogResult As CableSoft.BLL.Utility.RIAResult = CSLog.Summary(cmd, BLL.DataLog.OpType.Update, TableName, LogT.Tables(0), Integer.Parse(DateTime.Now.ToString("yyyyMMdd")))

        FNowDate = Date.Now
        Try
            If Not HavePK(EditMode, Account.Tables(FNewAccountTableName)) Then
                Return New RIAResult() With {.ErrorCode = -1, .ErrorMessage = "SO106 NO PKField", .ResultBoolean = False}
            End If
            Dim aAccountTB As DataTable = GetCorrectAccountTable(EditMode, Account)

            CSLog = New CableSoft.SO.BLL.DataLog.DataLog(Me.LoginInfo, Me.DAO)

            If DAO.Transaction IsNot Nothing Then

                trans = DAO.Transaction
            Else
                If cn IsNot Nothing AndAlso cn.State <> ConnectionState.Open Then
                    cn.ConnectionString = Me.LoginInfo.ConnectionString
                    cn.Open()
                End If
                trans = cn.BeginTransaction
                DAO.Transaction = trans
                blnAutoClose = True
            End If
            DAO.AutoCloseConn = False
            FNowDate = DateTime.Parse(DAO.ExecSclr(_DAL.GetSysDate))



            Using cmd As System.Data.Common.DbCommand = DAO._factory.CreateCommand()
                cmd.Connection = cn
                cmd.Transaction = trans
                'DAO.Transaction = trans
                Dim aWhere As String = String.Empty

                For i As Integer = 0 To aAccountTB.Rows.Count - 1
                    Dim aCitemStr As String = GetCitemStr(Account, cmd)
                    If Not String.IsNullOrEmpty(aCitemStr) Then
                        aAccountTB.Rows(i).Item("CitemStr") = aCitemStr
                    Else
                        aAccountTB.Rows(i).Item("CitemStr") = DBNull.Value
                    End If
                    aAccountTB.Rows(i).Item("UpdTime") = CableSoft.BLL.Utility.DateTimeUtility.GetDTString(FNowDate)
                    aAccountTB.Rows(i).Item("NewUpdTime") = FNowDate
                    If chkStopFlag(aAccountTB.Rows(i)) Then
                        aAccountTB.Rows(i).Item("ACHTNo") = DBNull.Value
                        aAccountTB.Rows(i).Item("ACHTDESC") = DBNull.Value
                    End If
                    cmd.Parameters.Clear()
                    '抓取SO002的InvSeqNo填入，避免提回出錯 By Kin 2012/09/06
                    Dim aInvSeqNo As String = Nothing
                    If DBNull.Value.Equals(aAccountTB.Rows(i).Item("INVSEQNO")) Then
                        aInvSeqNo = GetInvSeqNo(aAccountTB.Rows(i))
                    End If
                    If Not String.IsNullOrEmpty(aInvSeqNo) Then
                        aAccountTB.Rows(i).Item("INVSEQNO") = aInvSeqNo
                    End If
                    Select Case EditMode
                        Case CableSoft.BLL.Utility.EditMode.Append
                            If Not DAO.GetInsertOrUpdateCommand(CableSoft.Utility.DataAccess.UpdateMode.InsertRow, aAccountTB, "SO106", i, cmd, aWhere) Then
                                Return New RIAResult() With {.ErrorCode = -1, .ErrorMessage = Language.InsertSO106Error, .ResultBoolean = False}
                            End If
                        Case CableSoft.BLL.Utility.EditMode.Edit, CableSoft.BLL.Utility.EditMode.Delete
                            'If Account.Tables(FNewAccountTableName).Columns.Contains("RowId") Then
                            '    aWhere = String.Format("ROWID='{0}'",
                            '                       Account.Tables(FNewAccountTableName).Rows(i).Item("RowId"))
                            'Else
                            '    aWhere = String.Format("MasterId={0}",
                            '                          Account.Tables(FNewAccountTableName).Rows(i).Item("MasterId"))

                            'End If
                            If Account.Tables(FNewAccountTableName).Columns.Contains("RowId") Then
                                aWhere = String.Format("RowId='{0}'",
                                                   Account.Tables(FNewAccountTableName).Rows(i).Item("RowId"))
                            Else
                                aWhere = String.Format("MasterId={0}",
                                                      Account.Tables(FNewAccountTableName).Rows(i).Item("MasterId"))

                            End If
                            If EditMode = CableSoft.BLL.Utility.EditMode.Edit Then
                                If Not DAO.GetInsertOrUpdateCommand(CableSoft.Utility.DataAccess.UpdateMode.UpdateRow, aAccountTB, "SO106", i, cmd, aWhere) Then
                                    Return New RIAResult() With {.ErrorCode = -1, .ErrorMessage = Language.UpdateSO106Error, .ResultBoolean = False}
                                End If
                            Else
                                'cmd.CommandText = "DELETE SO106 WHERE " & aWhere
                                'cmd.CommandText = _DAL.DeleteSO106(aWhere)
                                DAO.ExecNqry(_DAL.DeleteSO106(aWhere))
                            End If

                    End Select
                    cmd.ExecuteNonQuery()
                    cmd.Parameters.Clear()
                    UpdSO003C(EditMode, Account, i, cmd)
                    cmd.Parameters.Clear()
                    ClearSO003C(EditMode, Account, i, False, cmd)
                    'UpdSO004(EditMode, Account, i, cmd)
                    'ClearSO004(EditMode, Account, i, False, cmd)
                    cmd.Parameters.Clear()
                    UpdSO003(EditMode, Account, i, cmd)
                    cmd.Parameters.Clear()
                    ClearSO003(EditMode, Account, i, False, cmd)
                    cmd.Parameters.Clear()
                    StopSO002A(EditMode, Account, i, cmd)
                    cmd.Parameters.Clear()
                    ChkSO002A(EditMode, Account.Tables(FNewAccountTableName).Rows(i), cmd)
                    cmd.Parameters.Clear()
                    'UpdSO033(EditMode, Account, i, cmd)
                    cmd.Parameters.Clear()
                    ProcessACH(EditMode, Account, i, cmd)
                    cmd.Parameters.Clear()
                    InsSO106Log(EditMode, Account, i, cmd)
                    cmd.Parameters.Clear()
                    Dim aResult As RIAResult = CSLog.SummaryExpansion(cmd, DataLog.OpType.Update, "SO106", Account.Tables(FNewAccountTableName), Int32.Parse(Integer.Parse(DateTime.Now.ToString("yyyyMMdd"))))
                    If Not aResult.ResultBoolean Then
                        Select Case aResult.ErrorCode
                            Case -5
                            Case -6
                                If blnAutoClose Then
                                    trans.Rollback()
                                    Return aResult
                                End If

                        End Select

                    End If
                Next

                aAccountTB.Dispose()
                If blnAutoClose Then
                    trans.Commit()
                End If

                Return New RIAResult() With {.ErrorCode = 0, .ErrorMessage = "OK",
                                             .ResultBoolean = True,
                                             .ResultXML = Account.Tables(FNewAccountTableName).Rows(0).Item("MasterId")}
            End Using

        Catch ex As Exception
            If (trans IsNot Nothing) AndAlso (blnAutoClose) Then
                trans.Rollback()
            End If
            Return New RIAResult() With {.ErrorCode = -1, .ErrorMessage = ex.Message, .ResultBoolean = False}
        Finally
            If blnAutoClose Then
                DAO.AutoCloseConn = True
                If trans IsNot Nothing Then
                    trans.Dispose()
                End If
                If cn IsNot Nothing Then
                    cn.Close()
                    cn.Dispose()
                End If

            End If
            If CSLog IsNot Nothing Then
                CSLog.Dispose()
            End If
        End Try
    End Function
    Private Function InsSO106Log(ByVal EditMode As EditMode, ByVal Account As DataSet,
                                 ByVal aRowIndex As Int32, ByRef cmd As DbCommand) As Boolean
        Try
            Dim aSO106Log As DataTable = CreateTableSchema("SO106LOG")
            Select Case EditMode
                Case CableSoft.BLL.Utility.EditMode.Append
                    Dim aRow As DataRow = aSO106Log.NewRow
                    aRow.Item("FuncType") = 3
                    For i As Int32 = 0 To Account.Tables(FNewAccountTableName).Columns.Count - 1
                        Select Case Account.Tables(FNewAccountTableName).Columns(i).ColumnName.ToUpper
                            'Case "rowid".ToUpper, "masterid".ToUpper, "compcode".ToUpper, "custid".ToUpper
                            Case "CTID".ToUpper, "masterid".ToUpper, "compcode".ToUpper, "custid".ToUpper, "ROWID".ToUpper
                            Case "UPDTIME".ToUpper
                                aRow.Item("UPDTIME") = CableSoft.BLL.Utility.DateTimeUtility.GetDTString(FNowDate)
                            Case "UPDEN".ToUpper
                                aRow.Item("UPDEN") = Me.LoginInfo.EntryName
                            Case "NEWUPDTIME".ToUpper
                                aRow.Item("NEWUPDTIME") = Account.Tables(FNewAccountTableName).Rows(aRowIndex).Item("NEWUPDTIME")
                            Case Else
                                If Account.Tables(FNewAccountTableName).Rows(aRowIndex).IsNull(
                                    Account.Tables(FNewAccountTableName).Columns(i).ColumnName) Then
                                    aRow.Item(Account.Tables(FNewAccountTableName).Columns(i).ColumnName & "B") = DBNull.Value
                                Else
                                    aRow.Item(Account.Tables(FNewAccountTableName).Columns(i).ColumnName & "B") =
                                        Account.Tables(FNewAccountTableName).Rows(aRowIndex).Item(Account.Tables(FNewAccountTableName).Columns(i).ColumnName)
                                End If
                        End Select
                    Next
                    aSO106Log.Rows.Add(aRow)
                    cmd.Parameters.Clear()
                    If Not DAO.GetInsertCommand(aSO106Log, "SO106LOG", 0, cmd) Then
                        Throw New Exception(Language.InsertSO106LogError)
                    End If
                Case CableSoft.BLL.Utility.EditMode.Edit
                    Dim blnChg As Boolean = False
                    For i As Int32 = 0 To Account.Tables(FNewAccountTableName).Columns.Count - 1
                        If Account.Tables(FNewAccountTableName).Columns(i).ColumnName.ToUpper <> "UPDTIME".ToUpper Then
                            If Not Account.Tables(FOldAccountTableName).Rows(0).Item(Account.Tables(FNewAccountTableName).Columns(i).ColumnName).Equals(
                                Account.Tables(FNewAccountTableName).Rows(aRowIndex).Item(Account.Tables(FNewAccountTableName).Columns(i).ColumnName)) Then

                                blnChg = True
                                Exit For
                            End If
                        End If
                    Next

                    If blnChg Then
                        Dim aRow As DataRow = aSO106Log.NewRow
                        aRow.Item("FuncType") = 1
                        For i As Int32 = 0 To Account.Tables(FNewAccountTableName).Columns.Count - 1
                            Select Case Account.Tables(FNewAccountTableName).Columns(i).ColumnName.ToUpper
                                'Case "rowid".ToUpper
                                Case "CTID".ToUpper, "ROWID".ToUpper
                                Case "NEWUPDTIME".ToUpper
                                    aRow.Item("NEWUPDTIME") = Account.Tables(FOldAccountTableName).Rows(aRowIndex).Item("NEWUPDTIME")
                                Case Else
                                    If aRow.Table.Columns.Contains(Account.Tables(FNewAccountTableName).Columns(i).ColumnName) Then
                                        If Account.Tables(FOldAccountTableName).Rows(0).IsNull(
                                       Account.Tables(FNewAccountTableName).Columns(i).ColumnName) Then
                                            aRow.Item(Account.Tables(FNewAccountTableName).Columns(i).ColumnName) = DBNull.Value
                                        Else
                                            aRow.Item(Account.Tables(FNewAccountTableName).Columns(i).ColumnName) =
                                                Account.Tables(FOldAccountTableName).Rows(0).Item(Account.Tables(FNewAccountTableName).Columns(i).ColumnName)

                                        End If
                                    End If

                            End Select
                        Next

                        For i As Int32 = 0 To Account.Tables(FNewAccountTableName).Columns.Count - 1
                            Select Case Account.Tables(FNewAccountTableName).Columns(i).ColumnName.ToUpper
                                'Case "rowid".ToUpper, "masterid".ToUpper, "compcode".ToUpper, "custid".ToUpper
                                Case "CTID".ToUpper, "masterid".ToUpper, "compcode".ToUpper, "custid".ToUpper, "ROWID".ToUpper
                                Case "UPDTIME".ToUpper
                                    aRow.Item("UPDTIME") = CableSoft.BLL.Utility.DateTimeUtility.GetDTString(FNowDate)
                                Case "UPDEN".ToUpper
                                    aRow.Item("UPDEN") = Me.LoginInfo.EntryName
                                Case "NEWUPDTIME".ToUpper
                                    aRow.Item("NEWUPDTIME") = Account.Tables(FNewAccountTableName).Rows(aRowIndex).Item("NEWUPDTIME")
                                Case Else
                                    If aRow.Table.Columns.Contains(Account.Tables(FNewAccountTableName).Columns(i).ColumnName & "B") Then
                                        If Account.Tables(FNewAccountTableName).Rows(aRowIndex).IsNull(
                                       Account.Tables(FNewAccountTableName).Columns(i).ColumnName) Then
                                            aRow.Item(Account.Tables(FNewAccountTableName).Columns(i).ColumnName & "B") = DBNull.Value
                                        Else
                                            aRow.Item(Account.Tables(FNewAccountTableName).Columns(i).ColumnName & "B") =
                                                Account.Tables(FNewAccountTableName).Rows(aRowIndex).Item(Account.Tables(FNewAccountTableName).Columns(i).ColumnName)
                                        End If
                                    End If
                            End Select
                        Next


                        aSO106Log.Rows.Add(aRow)
                        cmd.Parameters.Clear()
                        cmd.Parameters.Clear()
                        If Not DAO.GetInsertCommand(aSO106Log, "SO106LOG", 0, cmd) Then
                            Throw New Exception(Language.InsertSO106LogError)
                        End If
                        cmd.ExecuteNonQuery()
                    End If

                Case CableSoft.BLL.Utility.EditMode.Delete
                    Dim aRow As DataRow = aSO106Log.NewRow
                    aRow.Item("FuncType") = 2
                    For i As Int32 = 0 To Account.Tables(FNewAccountTableName).Columns.Count - 1
                        Select Case Account.Tables(FNewAccountTableName).Columns(i).ColumnName.ToUpper
                            'Case "rowid".ToUpper
                            'Case "NEWUPDTIME".ToUpper
                            Case "CTID".ToUpper, "ROWID".ToUpper
                            Case Else
                                If Account.Tables(FNewAccountTableName).Rows(aRowIndex).IsNull(Account.Tables(FNewAccountTableName).Columns(i).ColumnName) Then
                                    aRow.Item(Account.Tables(FNewAccountTableName).Columns(i).ColumnName) = DBNull.Value
                                Else
                                    aRow.Item(Account.Tables(FNewAccountTableName).Columns(i).ColumnName) =
                                        Account.Tables(FNewAccountTableName).Rows(aRowIndex).Item(Account.Tables(FNewAccountTableName).Columns(i).ColumnName)

                                End If
                        End Select
                    Next
                    aSO106Log.Rows.Add(aRow)
                    cmd.Parameters.Clear()
                    If Not DAO.GetInsertCommand(aSO106Log, "SO106LOG", 0, cmd) Then
                        Throw New Exception(Language.InsertSO106LogError)
                    End If
                    cmd.ExecuteNonQuery()

            End Select
            Return True
        Catch ex As Exception
            Throw
        End Try
    End Function
    Private Overloads Function GetAchCitemString(ByVal ACHTNo As String,
                                                ByVal Account As DataSet, ByVal aRowIndex As Integer) As Dictionary(Of String, String)
        Return GetAchCitemString(ACHTNo, Account, aRowIndex)
    End Function
    Private Overloads Function GetAchCitemString(ByVal ACHTNo As String,
                                                 ByVal Account As DataSet, ByVal aRowIndex As Integer, ByVal filterCustId As Boolean) As Dictionary(Of String, String)
        Dim cmd As DbCommand = DAO.GetConn.CreateCommand
        cmd.Transaction = DAO.Transaction
        cmd.Parameters.Clear()
        Dim ACHTCitem As Dictionary(Of String, String) = Nothing
        If newFlow Then
            ACHTCitem = GetAchCitem(ACHTNo, Account, aRowIndex, filterCustId, cmd)
        Else
            ACHTCitem = getOldAchCitem(ACHTNo, Integer.Parse(Account.Tables(FNewAccountTableName).Rows(aRowIndex).Item("CustId")), _
                                        Integer.Parse(Account.Tables(FNewAccountTableName).Rows(aRowIndex).Item("BankCode")), _
                                        Account.Tables(FNewAccountTableName).Rows(aRowIndex).Item("CitemStr"))
        End If

        Dim ret As New Dictionary(Of String, String)
        Dim CitemCode As String = Nothing
        Dim CitemName As String = Nothing
        If (ACHTCitem IsNot Nothing) AndAlso (ACHTCitem.Count > 0) Then
            For i As Integer = 0 To ACHTCitem.Count - 1
                If String.IsNullOrEmpty(CitemCode) Then
                    CitemCode = String.Format("{0}", ACHTCitem.Keys(i))
                Else
                    CitemCode = String.Format("{0},{1}", CitemCode, ACHTCitem.Keys(i))
                End If
                If String.IsNullOrEmpty(CitemName) Then
                    CitemName = String.Format("{0}", ACHTCitem.Values(i))
                Else
                    CitemName = String.Format("{0},{1}", CitemName, ACHTCitem.Values(i))
                End If
            Next
            ret.Add(CitemCode, CitemName)
        End If
        Return ret
    End Function
    Private Overloads Function GetAchCitem(ByVal aACHNo As String,
                                 ByVal Account As DataSet,
                                 ByVal aRowIndex As Integer) As Dictionary(Of String, String)
        Dim cmd As DbCommand = DAO.GetConn.CreateCommand
        cmd.Transaction = DAO.Transaction
        cmd.Parameters.Clear()
        Try
            Return GetAchCitem(aACHNo, Account, aRowIndex, cmd)
        Catch ex As Exception
            Throw
        Finally
            cmd.Dispose()
            cmd = Nothing
        End Try
    End Function
    Private Overloads Function GetAchCitem(ByVal aACHNo As String,
                                 ByVal Account As DataSet,
                                 ByVal aRowIndex As Int32,
                                 ByRef cmd As DbCommand) As Dictionary(Of String, String)
        If newFlow Then
            Return GetAchCitem(aACHNo, Account, aRowIndex, True, cmd)
        Else
            Return getOldAchCitem(aACHNo, _
                                  Integer.Parse(Account.Tables(FNewAccountTableName).Rows(aRowIndex).Item("CustId")), _
                                  Integer.Parse(Account.Tables(FNewAccountTableName).Rows(aRowIndex).Item("BankCode")), _
                                  Account.Tables(FNewAccountTableName).Rows(aRowIndex).Item("CitemStr"))
        End If

    End Function
    Private Function getOldAchCitem(ByVal aACHNo As String, ByVal CustId As Integer, ByVal BankCode As Integer, ByVal aSEQNo As String) As Dictionary(Of String, String)
        Dim starPos As Boolean = Integer.Parse(DAO.ExecSclr(_DAL.GetSystemPara, New Object() {LoginInfo.CompCode})) = 1
        Dim aRet As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
        Dim tbAchtNo As DataTable = DAO.ExecQry(_DAL.GetACHTNo(starPos))
        Dim tbBankCode As DataTable = DAO.ExecQry(_DAL.GetBankCodeByCode(starPos), New Object() {BankCode})
        Dim BillHeadFmt As String = "X"
        Try
            'For i As Integer = 0 To tbAchtNo.Rows.Count - 1
            '    If tbAchtNo.Rows(i).Item("ACHTNO") = aACHNo AndAlso tbAchtNo.Rows(i).Item("ACHTYPE") = Integer.Parse(tbAchtNo.Rows(0).Item("ACHTYPE")) Then
            '        BillHeadFmt = tbAchtNo.Rows(i).Item("BillHeadFmt")
            '    End If
            'Next
            BillHeadFmt = tbAchtNo.AsEnumerable.First(Function(rw As DataRow)
                                                          Return rw.Item("ACHTNO") = aACHNo AndAlso Integer.Parse(rw.Item("ACHTYPE")) = Integer.Parse(tbBankCode.Rows(0).Item("ACHTYPE"))
                                                      End Function).Item("BillHeadFmt")


            Dim tbCitemStr As DataTable = DAO.ExecQry(_DAL.getCD068A(aSEQNo), _
                                                      New Object() {BillHeadFmt, CustId, LoginInfo.CompCode})
            If tbCitemStr IsNot Nothing Then
                For Each rw As DataRow In tbCitemStr.Rows
                    If Not aRet.ContainsKey(rw.Item("CITEMCODE").ToString.ToUpper) Then
                        aRet.Add(rw.Item("CITEMCODE").ToString.ToUpper, rw.Item("CITEMNAME").ToString)
                    End If
                Next
            End If

            Return aRet
        Catch ex As Exception
            Throw ex
        Finally

        End Try
    End Function
    Private Overloads Function GetAchCitem(ByVal aACHNo As String,
                                 ByVal Account As DataSet,
                                 ByVal aRowIndex As Int32,
                                 ByVal filterCustId As Boolean,
                                 ByRef cmd As DbCommand) As Dictionary(Of String, String)

        'Select CitemCode From SO003 Where CustId = <ChooseFaci.CustId> And FaciSeqNo = <ChooseFaci.SeqNo> And StopFlag = 0
        Dim aRet As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
        Dim aServiceIds As String = "-99"
        Dim aFaciSeqNos As String = "'X'"
        Dim aProductCodes As String = "-99"
        Dim tbCitemStr As DataTable = Nothing
        Dim tbProduct As DataTable = Nothing
        Try
            For Each aRw As DataRow In Account.Tables(FChangeProductTableName).Rows
                If String.IsNullOrEmpty(aServiceIds) Then
                    aServiceIds = aRw.Item("ServiceId")
                Else
                    aServiceIds = aServiceIds & "," & aRw.Item("ServiceId")
                End If
            Next
            tbProduct = DAO.ExecQry(_DAL.GetAchCitem(Account.Tables(FNewAccountTableName).Rows(aRowIndex),
                                                    aServiceIds, aACHNo))
            If tbProduct IsNot Nothing Then
                For Each rw As DataRow In tbProduct.Rows
                    If Not DBNull.Value.Equals(rw.Item("ProductCode")) Then
                        aProductCodes = aProductCodes & "," & rw.Item("ProductCode")
                    End If
                    If Not DBNull.Value.Equals(rw.Item("FaciSeqNo")) Then
                        aFaciSeqNos = aFaciSeqNos & ",'" & rw.Item("FaciSeqNo") & "'"
                    End If
                Next
            End If


            If filterCustId Then
                tbCitemStr = DAO.ExecQry(_DAL.GetCitemCode(aFaciSeqNos, aProductCodes),
                                         New Object() {Account.Tables(FNewAccountTableName).Rows(aRowIndex).Item("CUSTID")})

            Else
                tbCitemStr = DAO.ExecQry(_DAL.GetNewCitemCode(aFaciSeqNos, aProductCodes),
                                         New Object() {Account.Tables(FDeclaredTableName).Rows(0).Item("SeqNo")})
            End If


            If tbCitemStr IsNot Nothing Then
                For Each rw As DataRow In tbCitemStr.Rows
                    If Not aRet.ContainsKey(rw.Item("CITEMCODE").ToString.ToUpper) Then
                        aRet.Add(rw.Item("CITEMCODE").ToString.ToUpper, rw.Item("CITEMNAME").ToString)
                    End If
                Next
            End If




            'If filterCustId Then
            '    cmd.CommandText = _DAL.GetCitemCode(Account.Tables(FNewAccountTableName).Rows(aRowIndex).Item("CUSTID"),
            '                                      aFaciSeqNos, aProductCodes)
            'Else
            '    cmd.CommandText = _DAL.GetNewCitemCode(Account.Tables(FDeclaredTableName).Rows(0).Item("SeqNo"), aFaciSeqNos, aProductCodes)
            'End If

            'Using dr As DbDataReader = cmd.ExecuteReader
            '    While dr.Read
            '        If Not aRet.ContainsKey(dr.Item("CITEMCODE").ToString.ToUpper) Then
            '            aRet.Add(dr.Item("CITEMCODE").ToString.ToUpper, dr.Item("CITEMNAME").ToString)
            '        End If

            '    End While
            'End Using
            Return aRet
        Catch ex As Exception
            Throw
        Finally
            If tbCitemStr IsNot Nothing Then
                tbCitemStr.Dispose()
                tbCitemStr = Nothing
            End If
            If tbProduct IsNot Nothing Then
                tbProduct.Dispose()
                tbProduct = Nothing
            End If
        End Try



    End Function
    Private Function CreateTableSchema(ByVal aTableName As String) As DataTable
        Dim aTableSchema As DataTable = Nothing
        Try
            'cmd.CommandText = _DAL.CreateTableSchema(aTableName)

            'Dim aTableSchema As DataTable = cmd.ExecuteReader.GetSchemaTable
            aTableSchema = DAO.ExecQry(_DAL.CreateTableSchema(aTableName))
            Dim aRetTB As New DataTable()
            For aIndex As Int32 = 0 To aTableSchema.Rows.Count - 1
                aRetTB.Columns.Add(aTableSchema.Rows(aIndex).Item("ColumnName").ToString,
                      aTableSchema.Rows(aIndex).Item("DataType"))
            Next
            Return aRetTB
        Catch ex As Exception
            Throw ex
        Finally
            If aTableSchema IsNot Nothing Then
                aTableSchema.Dispose()
                aTableSchema = Nothing
            End If
        End Try

    End Function
    Private Function UpdSO106A(ByVal EditMode As EditMode, ByVal Account As DataSet,
                               ByVal aRowIndex As Int32,
                               ByVal aAchtNo As String,
                               ByVal aCitem As Dictionary(Of String, String),
                               ByRef cmd As DbCommand) As Boolean

        Try
            Dim aCitemCode As String = String.Empty
            Dim aCitemName As String = String.Empty
            'Dim ACHTDESC As String = String.Empty
            Dim aWhere As String = "AuthorizeStatus = 4 AND MASTERID = " &
                Account.Tables(FNewAccountTableName).Rows(aRowIndex).Item("MASTERID") &
                    " AND ACHTNo = '" & aAchtNo.Trim("'") & "'"

            'cmd.CommandText = String.Format("Select ACHTDESC From CD068 Where ACHTNO='{0}'",
            '                              aAchtNo.Trim("'").ToString)
            'ACHTDESC = cmd.ExecuteScalar
            For i As Int32 = 0 To aCitem.Count - 1
                If String.IsNullOrEmpty(aCitemCode) Then
                    aCitemCode = aCitem.Keys(i)
                Else
                    aCitemCode = aCitemCode & "," & aCitem.Keys(i)
                End If
                If String.IsNullOrEmpty(aCitemName) Then
                    aCitemName = aCitem.Values(i)
                Else
                    aCitemName = aCitemName & "," & aCitem.Values(i)
                End If
            Next
            Dim aDTSO106A As DataTable = CreateTableSchema("SO106A")
            Dim adr As DataRow = aDTSO106A.NewRow


            If Not String.IsNullOrEmpty(aCitemCode) Then
                adr.Item("CitemCodeStr") = aCitemCode
                adr.Item("CitemNameStr") = aCitemName
            End If
            adr.Item("UpdEn") = Account.Tables(FNewAccountTableName).Rows(aRowIndex).Item("UpdEn")
            adr.Item("UpdTime") = Account.Tables(FNewAccountTableName).Rows(aRowIndex).Item("UpdTime")
            aDTSO106A.Rows.Add(adr)

            For Each col As DataColumn In aDTSO106A.Clone.Columns
                Select Case col.ColumnName.ToUpper
                    Case "CitemCodeStr".ToUpper, "CitemNameStr".ToUpper, "UpdEn".ToUpper, "UpdTime".ToUpper
                    Case Else
                        aDTSO106A.Columns.Remove(col.ColumnName)
                End Select
            Next
            cmd.Parameters.Clear()

            If Not DAO.GetUpdateCommand(aDTSO106A, "SO106A", 0, aWhere, cmd) Then
                Throw New Exception(Language.UpdateSO106AError)
            End If
            cmd.ExecuteNonQuery()
            'Update SO106A Set CitemCodeStr = <CitemCodeStr>,CitemNameStr = <CitemNameStr>, UpdEn=<使用者名稱>, UpdTime=<Account.UpdTime> Where MasterId = <Account.MasterId> And ACHTNo = <ACHTNO> And AuthorizeStatus = 4
            Return True
        Catch ex As Exception
            Throw ex
        Finally
            cmd.Parameters.Clear()
        End Try


    End Function
    Private Function DelSO106A(ByVal Account As DataSet, ByVal aRowIndex As Int32,
                               ByVal aAchtNo As String, ByRef cmd As DbCommand) As Boolean
        Try
            'Delete From SO106A Where MasterId = <Account.MasterId> And ACHTNO = <ACHTNO> And AuthorizeStatus = 4
            cmd.Parameters.Clear()
            'cmd.CommandText = String.Format("DELETE SO106A WHERE MasterId = {0} " &
            '        " AND ACHTNO ='{1}' AND AuthorizeStatus = 4 ", Account.Tables(FNewAccountTableName).Rows(aRowIndex).Item("MasterId"),
            '        aAchtNo.Trim("'"))
            'cmd.CommandText = _DAL.DeleteSO106A(Account.Tables(FNewAccountTableName).Rows(aRowIndex).Item("MasterId"), aAchtNo)
            'cmd.ExecuteNonQuery()
            DAO.ExecNqry(_DAL.DelSO106A, New Object() {Account.Tables(FNewAccountTableName).Rows(aRowIndex).Item("MasterId"),
                                                       aAchtNo.Trim("'")})
            Return True
        Catch ex As Exception
            Throw
        End Try
    End Function

    Private Function InsSO106A(ByVal EditMode As EditMode, ByVal Account As DataSet,
                               ByVal aRowIndex As Int32,
                               ByVal aAchtNo As String,
                               ByVal aCitem As Dictionary(Of String, String),
                               ByVal AuthStatus As AuthStatus,
                               ByRef cmd As DbCommand) As Boolean


        Try
            cmd.Parameters.Clear()
            'cmd.CommandText = String.Format("select rowid from so106 where masterid={0}",
            '                              Account.Tables(FNewAccountTableName).Rows(aRowIndex).Item("MasterId"))
            'Dim aRowId As String = cmd.ExecuteScalar
            Dim aRowId As String = DAO.ExecSclr(_DAL.GetSO106RowId, New Object() {Account.Tables(FNewAccountTableName).Rows(aRowIndex).Item("MasterId")})
            Dim aCitemCode As String = String.Empty
            Dim aCitemName As String = String.Empty
            Dim ACHTDESC As String = String.Empty
            'cmd.CommandText = String.Format("Select ACHTDESC From CD068 " & _
            '                                " Where ACHTNO='{0}'  " & _
            '                                " AND ACHTYPE =1 ",
            '                              aAchtNo.Trim("'").ToString)
            'ACHTDESC = cmd.ExecuteScalar
            ACHTDESC = DAO.ExecSclr(_DAL.GetACHTDESC, New Object() {aAchtNo.Trim("'").ToString})
            For i As Int32 = 0 To aCitem.Count - 1
                If String.IsNullOrEmpty(aCitemCode) Then
                    aCitemCode = aCitem.Keys(i)
                Else
                    aCitemCode = aCitemCode & "," & aCitem.Keys(i)
                End If
                If String.IsNullOrEmpty(aCitemName) Then
                    aCitemName = aCitem.Values(i)
                Else
                    aCitemName = aCitemName & "," & aCitem.Values(i)
                End If
            Next
            Dim aDTSO106A As DataTable = CreateTableSchema("SO106A")
            Dim aRwNew As DataRow = aDTSO106A.NewRow
            aRwNew.Item("MasterRowID") = aRowId

            If Not String.IsNullOrEmpty(aCitemCode) Then
                aRwNew.Item("CitemCodeStr") = aCitemCode
                aRwNew.Item("CitemNameStr") = aCitemName
            End If
            aRwNew.Item("ACHTNO") = aAchtNo.Trim("'")
            aRwNew.Item("UpdEn") = Account.Tables(FNewAccountTableName).Rows(aRowIndex).Item("UpdEn")
            aRwNew.Item("UpdTime") = Account.Tables(FNewAccountTableName).Rows(aRowIndex).Item("UpdTime")
            aRwNew.Item("CreateEn") = Account.Tables(FNewAccountTableName).Rows(aRowIndex).Item("UpdEn")
            aRwNew.Item("CreateTime") = FNowDate
            If AuthStatus = Customer.Account.Account.AuthStatus.Auth Then
                aRwNew.Item("RecordType") = 0
                aRwNew.Item("AuthorizeStatus") = 4
            Else
                aRwNew.Item("RecordType") = 1
                aRwNew.Item("AuthorizeStatus") = DBNull.Value
                aRwNew.Item("StopFlag") = 1
                If String.IsNullOrEmpty(GetStopDate(Account, aRowIndex)) Then
                    aRwNew.Item("StopDate") = FNowDate
                Else
                    aRwNew.Item("StopDate") = Date.Parse(GetStopDate(Account, aRowIndex))
                End If
            End If

            aRwNew.Item("MasterId") = Account.Tables(FNewAccountTableName).Rows(aRowIndex).Item("MasterId")
            aRwNew.Item("ACHDESC") = ACHTDESC
            aDTSO106A.Rows.Add(aRwNew)
            cmd.Parameters.Clear()
            If Not DAO.GetInsertCommand(aDTSO106A, "SO106A", 0, cmd) Then
                Throw New Exception(Language.InsertSO106AError)
            End If

            cmd.ExecuteNonQuery()
            Return True
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Private Function IsAddCancelAuth(ByVal aMasterId As String, ByVal aAchtNo As String, ByRef cmd As DbCommand) As Boolean
        Try
            'cmd.CommandText = String.Format("select count(*) from so106a " &
            '    " where AuthorizeStatus=2 and achtno='{0}'" &
            '    " and masterid={1}", aAchtNo.Trim("'"), aMasterId)
            'If Int32.Parse(cmd.ExecuteScalar) > 0 Then
            '    Return True
            'End If
            If Int32.Parse(DAO.ExecSclr(_DAL.IsAddCancelAuth, New Object() {aAchtNo.Trim("'"), aMasterId})) > 0 Then
                Return True
            End If
            Return False
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Private Function ProcessACHStatus(ByVal EditMode As EditMode,
                                ByVal Account As DataSet,
                                ByVal aRowIndex As Integer) As List(Of AchStatus)
        Return ProcessACHStatus(EditMode, Account, aRowIndex, True)
    End Function

    Private Function ProcessACHStatus(ByVal EditMode As EditMode,
                                ByVal Account As DataSet,
                                ByVal aRowIndex As Integer, ByVal filterCustId As Boolean) As List(Of AchStatus)
        Dim ret As New List(Of AchStatus)
        Dim tbSO106A As DataTable = Nothing
        Dim AryCitem As New Dictionary(Of String, String)
        Dim CitemCode As Object = DBNull.Value
        Dim CitemName As Object = DBNull.Value
        Try
            Dim lstOldACHTCode As List(Of String) = Nothing
            Dim lstOldACHTDesc As List(Of String) = Nothing
            Dim lstNewACHTCode As List(Of String) = Nothing
            Dim lstNewACHTDesc As List(Of String) = Nothing
            Dim rwSO106 As DataRow = Account.Tables(FNewAccountTableName).Rows(aRowIndex)
            If (Not DBNull.Value.Equals(rwSO106.Item("ACHTNo"))) AndAlso (Not String.IsNullOrEmpty(rwSO106.Item("ACHTNo").ToString)) Then
                lstNewACHTCode = rwSO106.Item("ACHTNo").ToString.Replace("'", "").Split(",").ToList
                lstNewACHTDesc = rwSO106.Item("ACHTDESC").ToString.Replace("'", "").Split(",").ToList
            Else
                lstNewACHTCode = New List(Of String)
                lstNewACHTDesc = New List(Of String)
            End If

            If (Account.Tables(FOldAch) IsNot Nothing) AndAlso (Account.Tables(FOldAch).Rows.Count > 0) AndAlso
                (Not String.IsNullOrEmpty(Account.Tables(FOldAch).Rows(0).Item(0).ToString)) Then
                lstOldACHTCode = Account.Tables(FOldAch).Rows(0).Item(0).ToString.Replace("'", "").Split(",").ToList
                lstOldACHTDesc = Account.Tables(FOldAch).Rows(0).Item(1).ToString.Replace("'", "").Split(",").ToList
            Else
                lstOldACHTCode = New List(Of String)
                lstOldACHTDesc = New List(Of String)
            End If

            If Integer.Parse("0" & rwSO106.Item("StopFlag")) = 1 Then
                lstNewACHTCode.Clear()
                lstNewACHTDesc.Clear()
            End If
            If EditMode = CableSoft.BLL.Utility.EditMode.Append Then
                If (lstNewACHTCode Is Nothing) OrElse (lstNewACHTCode.Count = 0) Then
                    Return ret
                Else
                    For i As Integer = 0 To lstNewACHTCode.Count - 1
                        Dim retStatus As New AchStatus
                        AryCitem.Clear()
                        AryCitem = GetAchCitemString(lstNewACHTCode.Item(i), Account, aRowIndex, filterCustId)
                        retStatus.ACHTNo = lstNewACHTCode.Item(i)
                        retStatus.ACHTDesc = lstNewACHTDesc.Item(i)
                        retStatus.UpdateType = AchUpdateType.AddAuthorize
                        If (AryCitem IsNot Nothing) AndAlso (AryCitem.Count > 0) Then
                            retStatus.CitemCode = AryCitem.Keys(0)
                            retStatus.CitemName = AryCitem.Values(0)
                        End If
                        ret.Add(retStatus)
                    Next
                End If
            Else
                '新跟舊都沒有ACH
                If (lstNewACHTCode Is Nothing) OrElse (lstNewACHTCode.Count = 0) Then
                    If (lstOldACHTCode Is Nothing) OrElse (lstOldACHTCode.Count = 0) Then
                        Return ret
                    End If
                End If
                Dim AddAch As New Dictionary(Of String, String)
                Dim DelAch As New Dictionary(Of String, String)
                Dim EditAch As New Dictionary(Of String, String)
                '畫面上有,但舊資料沒有
                If (lstNewACHTCode IsNot Nothing) AndAlso (lstNewACHTCode.Count > 0) Then
                    For i As Integer = 0 To lstNewACHTDesc.Count - 1
                        If (lstOldACHTDesc Is Nothing) OrElse (Not lstOldACHTDesc.Contains(lstNewACHTDesc.Item(i))) Then
                            AddAch.Add(lstNewACHTDesc.Item(i), lstNewACHTCode.Item(i))
                        End If
                    Next
                End If
                '畫面上沒有,但舊資料有
                If (lstOldACHTDesc IsNot Nothing) AndAlso (lstOldACHTDesc.Count > 0) Then
                    For i As Integer = 0 To lstOldACHTDesc.Count - 1
                        If (lstNewACHTDesc Is Nothing) OrElse (Not lstNewACHTDesc.Contains(lstOldACHTDesc.Item(i))) Then
                            DelAch.Add(lstOldACHTDesc.Item(i), lstOldACHTCode.Item(i))
                        End If
                    Next
                End If
                '畫面上有,舊資料也有
                If (lstNewACHTDesc IsNot Nothing) AndAlso (lstOldACHTDesc IsNot Nothing) Then
                    For i As Integer = 0 To lstNewACHTDesc.Count - 1
                        If lstOldACHTDesc.Contains(lstNewACHTDesc.Item(i)) Then
                            EditAch.Add(lstNewACHTDesc.Item(i), lstNewACHTCode.Item(i))
                        End If
                    Next
                End If


                For i As Integer = 0 To AddAch.Count - 1
                    AryCitem.Clear()
                    CitemCode = DBNull.Value
                    CitemName = DBNull.Value
                    AryCitem = GetAchCitemString(AddAch.Values(i), Account, aRowIndex, filterCustId)
                    If (AryCitem IsNot Nothing) AndAlso (AryCitem.Count > 0) Then
                        CitemCode = AryCitem.Keys(0)
                        CitemName = AryCitem.Values(0)
                    End If


                    ret.Add(New AchStatus With {.UpdateType = AchUpdateType.AddAuthorize,
                                                .ACHTNo = AddAch.Values(i), .ACHTDesc = AddAch.Keys(i),
                                                .CitemCode = CitemCode, .CitemName = CitemName})

                Next

                For i As Integer = 0 To DelAch.Count - 1
                    AryCitem.Clear()
                    CitemCode = DBNull.Value
                    CitemName = DBNull.Value
                    AryCitem = GetAchCitemString(DelAch.Values(i), Account, aRowIndex, filterCustId)
                    If (AryCitem IsNot Nothing) AndAlso (AryCitem.Count > 0) Then
                        CitemCode = AryCitem.Keys(0)
                        CitemName = AryCitem.Values(0)
                    End If

                    ret.Add(New AchStatus With {.UpdateType = AchUpdateType.CancelAuthorize,
                                                .ACHTNo = DelAch.Values(i), .ACHTDesc = DelAch.Keys(i),
                                                .CitemCode = CitemCode, .CitemName = CitemName})
                Next

                For i As Integer = 0 To EditAch.Count - 1
                    AryCitem.Clear()
                    CitemCode = DBNull.Value
                    CitemName = DBNull.Value
                    AryCitem = GetAchCitemString(EditAch.Values(i), Account, aRowIndex, filterCustId)
                    If (AryCitem IsNot Nothing) AndAlso (AryCitem.Count > 0) Then
                        CitemCode = AryCitem.Keys(0)
                        CitemName = AryCitem.Values(0)
                    End If

                    ret.Add(New AchStatus With {.UpdateType = AchUpdateType.ChangeCitem,
                                                .ACHTNo = EditAch.Values(i), .ACHTDesc = EditAch.Keys(i),
                                                .CitemCode = CitemCode, .CitemName = CitemName})
                Next


            End If
        Catch ex As Exception
            Throw
        Finally
            If tbSO106A IsNot Nothing Then
                tbSO106A.Dispose()
                tbSO106A = Nothing
            End If
        End Try
        Return ret

    End Function
    Private Function ProcessACH(ByVal EditMode As EditMode,
                                ByVal Account As DataSet,
                                ByVal aRowIndex As Int32,
                                ByRef cmd As DbCommand) As Boolean
        Return ProcessACH(EditMode, Account, aRowIndex, True, cmd)
    End Function

    Private Function ProcessACH(ByVal EditMode As EditMode,
                                ByVal Account As DataSet,
                                ByVal aRowIndex As Int32,
                                ByVal filterCustId As Boolean,
                                ByRef cmd As DbCommand) As Boolean
        Dim tbSO106A As DataTable = Nothing

        Try
            Dim lstStatus As List(Of AchStatus) = ProcessACHStatus(EditMode, Account, aRowIndex, filterCustId)
            Dim aRowId As String =
               DAO.ExecSclr(_DAL.GetSO106RowId, New Object() {Account.Tables(FNewAccountTableName).Rows(aRowIndex).Item("MasterId")})
            tbSO106A = DAO.ExecQry(_DAL.GetSO106A(), New Object() {
                                   Account.Tables(FNewAccountTableName).Rows(aRowIndex).Item("MasterId")})
            If (lstStatus IsNot Nothing) AndAlso (lstStatus.Count > 0) Then
                For Each status As AchStatus In lstStatus
                    Select Case status.UpdateType
                        Case AchUpdateType.AddAuthorize
                            DAO.ExecNqry(_DAL.AddAuthorize,
                                         New Object() {aRowId, status.ACHTNo,
                                                       status.CitemCode, status.CitemName,
                                                       LoginInfo.EntryName, CableSoft.BLL.Utility.DateTimeUtility.GetDTString(FNowDate), FNowDate, LoginInfo.EntryName, 0,
                                                       4, status.ACHTDesc, Account.Tables(FNewAccountTableName).Rows(aRowIndex).Item("MasterId"),
                                                       0, DBNull.Value})

                        Case AchUpdateType.CancelAuthorize
                            If Integer.Parse(DAO.ExecNqry(_DAL.DelWaitAuthorize(status.CitemCode, status.CitemName),
                                                          New Object() {Account.Tables(FNewAccountTableName).Rows(aRowIndex).Item("MasterId"),
                                                                              status.ACHTNo, status.ACHTDesc})) = 0 Then


                                DAO.ExecNqry(_DAL.AddAuthorize,
                                         New Object() {aRowId, status.ACHTNo,
                                                       status.CitemCode, status.CitemName,
                                                       LoginInfo.EntryName, CableSoft.BLL.Utility.DateTimeUtility.GetDTString(FNowDate), FNowDate, LoginInfo.EntryName, 1,
                                                       4, status.ACHTDesc, Account.Tables(FNewAccountTableName).Rows(aRowIndex).Item("MasterId"), 1,
                                                       FNowDate})


                            End If
                        Case AchUpdateType.ChangeCitem
                            Using tb1 As DataTable = DAO.ExecQry(_DAL.QueryAccountDetail(), New Object() {Account.Tables(FNewAccountTableName).Rows(aRowIndex).Item("MasterId")})
                                For Each rw As DataRow In tb1.Rows
                                    If (Integer.Parse("0" & rw.Item("StopFlag").ToString) = 0) AndAlso
                                       (rw.Item("ACHTNO").ToString = status.ACHTNo) AndAlso (rw.Item("ACHDesc").ToString = status.ACHTDesc) AndAlso
                                       (rw.Item("CitemNameStr").ToString <> status.CitemName.ToString) Then
                                        DAO.ExecNqry(_DAL.UpdAuthorize(), New Object() {
                                                                 status.CitemCode, status.CitemName, LoginInfo.EntryName,
                                                              CableSoft.BLL.Utility.DateTimeUtility.GetDTString(FNowDate),
                                                             rw.Item("ctid")})
                                    End If
                                Next
                            End Using

                            'Using dr As DbDataReader = DAO.ExecDtRdr(_DAL.QueryAccountDetail(), New Object() {Account.Tables(FNewAccountTableName).Rows(aRowIndex).Item("MasterId")})
                            '    While dr.Read
                            '        If (Integer.Parse("0" & dr.Item("StopFlag").ToString) = 0) AndAlso
                            '            (dr.Item("ACHTNO").ToString = status.ACHTNo) AndAlso (dr.Item("ACHDesc").ToString = status.ACHTDesc) AndAlso
                            '            (dr.Item("CitemNameStr").ToString <> status.CitemName.ToString) Then

                            '            DAO.ExecNqry(_DAL.UpdAuthorize(), New Object() {
                            '                                     status.CitemCode, status.CitemName, LoginInfo.EntryName,
                            '                                  CableSoft.BLL.Utility.DateTimeUtility.GetDTString(FNowDate),
                            '                                 dr.Item("ctid")})

                            '        End If


                            '    End While
                            'End Using



                    End Select
                Next
            End If

            Return True
            If (Account.Tables(FNewAccountTableName).Rows(aRowIndex).IsNull("ACHTNo")) OrElse
               (String.IsNullOrEmpty(Account.Tables(FNewAccountTableName).Rows(aRowIndex).Item("ACHTNO"))) Then
                Return True
            End If

            If chkStopFlag(Account.Tables(FNewAccountTableName).Rows(aRowIndex)) Then Return True

            Dim aryACH As String() = Account.Tables(FNewAccountTableName).Rows(aRowIndex).Item("ACHTNO").ToString.Replace("'", "").Split(",")
            Dim aryOLDACH As String()
            If EditMode = CableSoft.BLL.Utility.EditMode.Append Then
                aryOLDACH = aryACH
            Else
                aryOLDACH = Account.Tables(FOldAch).Rows(0).Item(0).ToString.Replace("'", "").Split(",")
            End If


            For Each s As String In aryACH
                Dim objCitem As Dictionary(Of String, String) = GetAchCitem(s, Account, aRowIndex, filterCustId, cmd)
                Select Case EditMode
                    Case CableSoft.BLL.Utility.EditMode.Append
                        InsSO106A(EditMode, Account, aRowIndex, s, objCitem, AuthStatus.Auth, cmd)
                    Case CableSoft.BLL.Utility.EditMode.Edit
                        If aryOLDACH.Contains(s) Then
                            UpdSO106A(EditMode, Account, aRowIndex, s, objCitem, cmd)
                        Else
                            InsSO106A(EditMode, Account, aRowIndex, s, objCitem, AuthStatus.Auth, cmd)
                        End If
                End Select
            Next

            For Each s As String In aryOLDACH
                If EditMode = CableSoft.BLL.Utility.EditMode.Edit Then
                    If Not aryACH.Contains(s) Then
                        DelSO106A(Account, aRowIndex, s, cmd)
                        If IsAddCancelAuth(
                            Account.Tables(FNewAccountTableName).Rows(aRowIndex).Item("MasterId"), s, cmd) Then
                            Dim objCitem As Dictionary(Of String, String) = GetAchCitem(s, Account, aRowIndex, cmd)
                            InsSO106A(EditMode, Account, aRowIndex, s, objCitem, AuthStatus.Cancel, cmd)
                        End If
                    End If
                End If
            Next

            Return True
        Catch ex As Exception
            Throw ex
        Finally
            If tbSO106A IsNot Nothing Then
                tbSO106A.Dispose()
                tbSO106A = Nothing
            End If
        End Try
    End Function
    Private Function UpdSO003(ByVal EditMode As EditMode,
                             ByVal Account As DataSet, ByVal aRowIndex As Integer,
                             ByRef cmd As DbCommand) As Boolean

        Return UpdSO003(EditMode, Account, aRowIndex, True, cmd)
    End Function
    Private Function stopOldData(ByVal Account As DataSet) As Boolean
        If newFlow Then Return True
        Dim citemCode As Object = DBNull.Value
        Dim strBillNos As String = Nothing
        Dim strServiceType As String = Nothing
        Try
            If Not DBNull.Value.Equals(Account.Tables(FNewAccountTableName).Rows(0).Item("StopFlag")) AndAlso
                Integer.Parse(Account.Tables(FNewAccountTableName).Rows(0).Item("StopFlag")) = 1 Then
                Dim PtCode As Integer = 1
                Dim ptName As String = Language.ptCash
                Dim CmCode As Integer = 1
                Dim CmName As String = ""
                Dim Uccode As Object = DBNull.Value
                Dim UcName As Object = DBNull.Value
                Using tb As DataTable = DAO.ExecQry(_DAL.GetDefPTCode)
                    If tb IsNot Nothing Then
                        PtCode = tb.Rows(0).Item("CODENO")
                        ptName = tb.Rows(0).Item("Description")
                        tb.Dispose()
                    End If
                End Using

                Using tb As DataTable = DAO.ExecQry(_DAL.GetDefCMCode(LoginInfo, String.Empty))
                    If tb IsNot Nothing Then
                        CmCode = tb.Rows(0).Item("CODENO")
                        CmName = tb.Rows(0).Item("Description")
                        tb.Dispose()
                    End If

                End Using

                With Account.Tables(FNewAccountTableName).Rows(0)
                    If DBNull.Value.Equals(.Item("CitemStr")) Then Return True
                    For Each seq As String In Split(.Item("CitemStr").ToString, ",")
                        DAO.ExecNqry(_DAL.StopSO003BySeq, New Object() {CmCode, CmName, PtCode, ptName, .Item("CustId"), _
                                                                            .Item("AccountId"), LoginInfo.CompCode, .Item("BankCode"), Integer.Parse(seq.Replace("'", ""))})
                        If Integer.Parse(DAO.ExecSclr(_DAL.chkSO003(seq), New Object() {.Item("CustId"), .Item("AccountId"), _
                                                                       LoginInfo.CompCode, Integer.Parse(.Item("MasterId"))})) = 0 Then

                            citemCode = DAO.ExecSclr(_DAL.QuerySO003CitemBySeq, New Object() {.Item("CustId"), .Item("AccountId"), _
                                                                                              LoginInfo.CompCode, Integer.Parse(seq.Replace("'", "")), .Item("BankCode")})

                            If citemCode Then
                                Using tbSO033 As DataTable = DAO.ExecQry(_DAL.QuerySO033, New Object() {.Item("CustId"), _
                                                                                                        .Item("AccountId"), LoginInfo.CompCode, .Item("BankCode"), Integer.Parse(citemCode)})
                                    DAO.ExecNqry(_DAL.stopOldSO033, New Object() {CmCode, CmName, PtCode, ptName, .Item("CustId"), _
                                                                                  .Item("AccountId"), LoginInfo.CompCode, .Item("BankCode"), Integer.Parse(citemCode)})

                                    If tbSO033.Rows.Count > 0 Then
                                        strBillNos = tbSO033.Rows(0).Item("BillNo")
                                        strServiceType = tbSO033.Rows(0).Item("ServiceType")
                                        Using tbUccode As DataTable = DAO.ExecQry(_DAL.QueryUccode, New Object() {strServiceType, LoginInfo.CompCode})
                                            If tbUccode.Rows.Count > 0 Then
                                                Uccode = tbUccode.Rows(0).Item("CodeNo")
                                                UcName = tbUccode.Rows(0).Item("Description")
                                            End If
                                        End Using
                                        DAO.ExecNqry(_DAL.updOldSO033Uccode, New Object() {CmCode, CmName, _
                                                                                          PtCode, ptName, Uccode, UcName, _
                                                                                          .Item("CustId"), LoginInfo.CompCode, strBillNos})

                                    End If

                                End Using
                            End If
                        End If

                    Next
                End With
            End If

        Catch ex As Exception
            Throw ex
        Finally

        End Try

        Return True
    End Function
    Private Function UpdNonePeriod(ByVal EditMode As EditMode,
                              ByVal Account As DataSet, ByVal aRowIndex As Integer, ByVal tbAccount As DataTable, ByVal SeqNo As String,
                              ByRef cmd As DbCommand)
        Try

            If chkStopFlag(Account.Tables(FNewAccountTableName).Rows(0)) AndAlso newFlow Then
                Return True
            End If
            If String.IsNullOrEmpty(SeqNo) Then
                Return True
            End If
            Dim a106Rw As DataRow = Account.Tables(FNewAccountTableName).Rows(aRowIndex)
            If newFlow Then
                DAO.ExecNqry(String.Format(_DAL.UpdNewNonePeriod(SeqNo),
                                              a106Rw.Item("AccountId").ToString,
                                              a106Rw.Item("BankCode"), a106Rw.Item("BankName"),
                                              a106Rw.Item("CMCode"), a106Rw.Item("CMName"),
                                              a106Rw.Item("PTCode"), a106Rw.Item("PTName"),
                                               a106Rw.Item("UpdEn"),
                                               a106Rw.Item("UpdTime"),
                                               CType(a106Rw.Item("NewUpdTime"), Date).ToString("yyyyMMddHHmmss"),
                                               Account.Tables(FDeclaredTableName).Rows(0).Item("SEQNO")))
            Else
                If chkSnactionDate(Account.Tables(FNewAccountTableName).Rows(aRowIndex)) Then
                    If EditMode = CableSoft.BLL.Utility.EditMode.Append Then
                        tbAccount = Account.Tables(FNewAccountTableName).Copy
                    End If
                    DAO.ExecNqry(String.Format(_DAL.UpdOldNonePeriod(SeqNo, a106Rw.Item("CustID")),
                                             a106Rw.Item("AccountId").ToString,
                                             a106Rw.Item("BankCode"), a106Rw.Item("BankName"),
                                             a106Rw.Item("CMCode"), a106Rw.Item("CMName"),
                                             a106Rw.Item("PTCode"), a106Rw.Item("PTName"),
                                              a106Rw.Item("UpdEn"),
                                              a106Rw.Item("UpdTime"), SO138_InvSeqNo,
                                              CType(a106Rw.Item("NewUpdTime"), Date).ToString("yyyyMMddHHmmss")))
                End If

            End If


            'cmd.CommandText = String.Format(_DAL.UpdNewNonePeriod(SeqNo),
            '                                  a106Rw.Item("AccountId").ToString,
            '                                  a106Rw.Item("BankCode"), a106Rw.Item("BankName"),
            '                                  a106Rw.Item("CMCode"), a106Rw.Item("CMName"),
            '                                  a106Rw.Item("PTCode"), a106Rw.Item("PTName"),
            '                                   a106Rw.Item("UpdEn"),
            '                                   a106Rw.Item("UpdTime"),
            '                                   CType(a106Rw.Item("NewUpdTime"), Date).ToString("yyyyMMddHHmmss"),
            '                                   Account.Tables(FDeclaredTableName).Rows(0).Item("SEQNO"))
            'cmd.ExecuteNonQuery()
            Return True

        Catch ex As Exception
            Throw
        End Try

    End Function

    ''' <summary>
    ''' 更新SO003
    ''' </summary>
    ''' <param name="EditMode">狀態</param>
    ''' <param name="Account">DataSet</param>
    ''' <param name="aRowIndex">Account RowIndex</param>
    ''' <param name="cmd">DBCommand</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function UpdSO003(ByVal EditMode As EditMode,
                              ByVal Account As DataSet, ByVal aRowIndex As Integer, ByVal filterCustId As Boolean,
                              ByRef cmd As DbCommand) As Boolean
        Dim tbSO003C As DataTable = Nothing
        Try

            If chkStopFlag(Account.Tables(FNewAccountTableName).Rows(0)) Then
                Return True
            End If
            If Not newFlow Then Return True
            Dim a106Rw As DataRow = Account.Tables(FNewAccountTableName).Rows(aRowIndex)
            Dim aServiceIds As String = "-99"
            Dim aProdcutCodes As String = "-99"
            Dim aFaciSeqNos As String = "'X'"
            If EditMode = CableSoft.BLL.Utility.EditMode.Append Then
            Else

            End If
            If chkSnactionDate(Account.Tables(FNewAccountTableName).Rows(0)) Then
                For Each aRw As DataRow In Account.Tables(FChangeProductTableName).Rows
                    aServiceIds = aServiceIds & "," & aRw.Item("ServiceId") & ""
                Next
            End If
            'cmd.Parameters.Clear()


            If filterCustId Then
                '                cmd.CommandText = String.Format(_DAL.GetSO003C, aServiceIds, a106Rw.Item("CUSTID"))
                tbSO003C = DAO.ExecQry(String.Format(_DAL.GetSO003C, aServiceIds, a106Rw.Item("CUSTID")))
            Else
                'cmd.CommandText = String.Format(_DAL.GetNewSO003C, aServiceIds, Account.Tables(FDeclaredTableName).Rows(0).Item("SEQNO"))
                tbSO003C = DAO.ExecQry(String.Format(_DAL.GetNewSO003C, aServiceIds, Account.Tables(FDeclaredTableName).Rows(0).Item("SEQNO")))
            End If
            If tbSO003C IsNot Nothing Then
                For Each rw As DataRow In tbSO003C.Rows
                    If Not DBNull.Value.Equals(rw.Item("ProductCode")) Then
                        aProdcutCodes = aProdcutCodes & "," & rw.Item("ProductCode")
                    End If
                    If Not DBNull.Value.Equals(rw.Item("FaciSeqNo")) Then
                        aFaciSeqNos = aFaciSeqNos & ",'" & rw.Item("FaciSeqNo") & "'"
                    End If
                Next
            End If



            If filterCustId Then

                DAO.ExecNqry(String.Format(_DAL.UpdSO003,
                                                a106Rw.Item("AccountId").ToString,
                                                a106Rw.Item("BankCode"), a106Rw.Item("BankName"),
                                                a106Rw.Item("CMCode"), a106Rw.Item("CMName"),
                                                a106Rw.Item("PTCode"), a106Rw.Item("PTName"),
                                                aFaciSeqNos, a106Rw.Item("CUSTID"), aProdcutCodes))
            Else

                DAO.ExecNqry(String.Format(_DAL.UpdNewSO003,
                                               a106Rw.Item("AccountId").ToString,
                                               a106Rw.Item("BankCode"), a106Rw.Item("BankName"),
                                               a106Rw.Item("CMCode"), a106Rw.Item("CMName"),
                                               a106Rw.Item("PTCode"), a106Rw.Item("PTName"),
                                                a106Rw.Item("UpdEn"),
                                                a106Rw.Item("UpdTime"),
                                                CType(a106Rw.Item("NewUpdTime"), Date).ToString("yyyyMMddHHmmss"),
                                               aFaciSeqNos, aProdcutCodes, Account.Tables(FDeclaredTableName).Rows(0).Item("SEQNO")))
            End If
            'cmd.ExecuteNonQuery()
            Return True
        Catch ex As Exception
            Throw
        Finally
            If tbSO003C IsNot Nothing Then
                tbSO003C.Dispose()
                tbSO003C = Nothing
            End If
        End Try
    End Function
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="aRow"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function chkSnactionDate(ByVal aRow As DataRow) As Boolean
        Try
            If aRow.IsNull("SnactionDate") Then
                Return False
            End If
            If String.IsNullOrEmpty(aRow.Item("SnactionDate").ToString) Then
                Return False
            End If
            Dim aDate As Date = Date.Now
            If Not Date.TryParse(aRow.Item("SnactionDate").ToString, aDate) Then
                Return False
            End If
            Return True
        Catch ex As Exception
            Throw
        End Try
    End Function
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="aRow"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function chkStopFlag(ByVal aRow As DataRow) As Boolean
        Try
            If (Not aRow.IsNull("StopFlag")) AndAlso
                (Integer.Parse(aRow.Item("StopFlag")) = 1) Then
                Return True
            End If
            Return False
        Catch ex As Exception
            Throw
        End Try
    End Function
    Private Function GetStopDate(ByVal Account As DataSet, ByVal aRowIndex As Int32) As String
        Try
            If (Account.Tables(FNewAccountTableName).Rows(aRowIndex).IsNull("StopDate")) OrElse
                (String.IsNullOrEmpty(Account.Tables(FNewAccountTableName).Rows(aRowIndex).Item("StopDate"))) Then
                Return String.Empty
            End If
            Return Date.Parse(Account.Tables(FNewAccountTableName).Rows(aRowIndex).Item("StopDate")).ToString
        Catch ex As Exception
            Return String.Empty
        End Try
    End Function
    Private Function ClearSO003(ByVal EditMode As EditMode,
                                ByVal Account As DataSet, ByVal aRowIndex As Integer,
                                ByVal VoidData As Boolean,
                                ByRef cmd As DbCommand) As Boolean
        Return ClearSO003(EditMode, Account, aRowIndex, VoidData, True, cmd)
    End Function
    Private Function ClearOldSO003(ByVal EditMode As EditMode, ByVal Account As DataSet, ByVal tbAccount As DataTable, ByVal SEQNO As String, ByRef cmd As DbCommand) As Boolean
        If newFlow Then Return True
        If Not chkStopFlag(Account.Tables(FNewAccountTableName).Rows(0)) Then
            Return True
        End If
        Return ClearNonePeriod(EditMode, Account, tbAccount, SEQNO, cmd)
    End Function
    Private Function ClearNonePeriod(ByVal EditMode As EditMode, ByVal Account As DataSet, ByVal tbAccount As DataTable, ByVal SEQNO As String, ByRef cmd As DbCommand) As Boolean
        Dim aPTCode As Integer = -1
        Dim aPTName As String = String.Empty
        Dim aCMCode As Integer = -1
        Dim aCMName As String = String.Empty
        Dim aCustId As String = "-1"

        If String.IsNullOrEmpty(SEQNO) AndAlso newFlow Then Return True
        If EditMode = CableSoft.BLL.Utility.EditMode.Append Then Return True
        Try
            Using tb As DataTable = DAO.ExecQry(_DAL.GetPTCode)
                If tb IsNot Nothing Then
                    If tb.Rows.Count > 0 Then
                        aPTCode = tb.Rows(0).Item("CODENO")
                        aPTName = tb.Rows(0).Item("Description")
                    End If
                End If

            End Using
            'cmd.CommandText = _DAL.GetPTCode
            'Using dr As DbDataReader = cmd.ExecuteReader()
            '    dr.Read()
            '    aPTCode = dr.Item("CODENO")
            '    aPTName = dr.Item("Description")
            'End Using
            Using tb As DataTable = DAO.ExecQry(_DAL.GetDefCMCode(LoginInfo, String.Empty))
                If tb IsNot Nothing Then
                    If tb.Rows.Count > 0 Then
                        aCMCode = tb.Rows(0).Item("CODENO")
                        aCMName = tb.Rows(0).Item("Description")
                    End If
                End If

            End Using

            'cmd.CommandText = _DAL.GetDefCMCode(LoginInfo, String.Empty)

            'Using dr As DbDataReader = cmd.ExecuteReader
            '    dr.Read()
            '    aCMCode = dr.Item("CODENO")
            '    aCMName = dr.Item("Description")
            'End Using
            If newFlow Then
                DAO.ExecNqry(String.Format(_DAL.ClearNoneSO003(SEQNO),
                                               aCMCode, aCMName, aPTCode, aPTName,
                                          Account.Tables(FNewAccountTableName).Rows(0).Item("UpdEn"),
                                           Account.Tables(FNewAccountTableName).Rows(0).Item("UpdTime"),
                                          CType(Account.Tables(FNewAccountTableName).Rows(0).Item("NewUpdTime"), Date).ToString("yyyyMMddHHmmss"),
                                               Account.Tables(FDeclaredTableName).Rows(0).Item("SEQNO")))
            Else
                DAO.ExecNqry(String.Format(_DAL.ClearOldNoneSO003(SEQNO, Account.Tables(FNewAccountTableName).Rows(0).Item("CustId")),
                                               aCMCode, aCMName, aPTCode, aPTName,
                                          Account.Tables(FNewAccountTableName).Rows(0).Item("UpdEn"),
                                           Account.Tables(FNewAccountTableName).Rows(0).Item("UpdTime"),
                                          CType(Account.Tables(FNewAccountTableName).Rows(0).Item("NewUpdTime"), Date).ToString("yyyyMMddHHmmss"), _
                                        Integer.Parse(tbAccount.Rows(0).Item("BankCode")), _
                                        Me.LoginInfo.CompCode, tbAccount.Rows(0).Item("AccountId")))
            End If


            'cmd.CommandText = String.Format(_DAL.ClearNoneSO003(SEQNO),
            '                                    aCMCode, aCMName, aPTCode, aPTName,
            '                               Account.Tables(FNewAccountTableName).Rows(0).Item("UpdEn"),
            '                                Account.Tables(FNewAccountTableName).Rows(0).Item("UpdTime"),
            '                               CType(Account.Tables(FNewAccountTableName).Rows(0).Item("NewUpdTime"), Date).ToString("yyyyMMddHHmmss"),
            '                                    Account.Tables(FDeclaredTableName).Rows(0).Item("SEQNO"))
            'cmd.ExecuteNonQuery()
        Catch ex As Exception
            Throw
        End Try
        Return True
    End Function
    Private Function ClearSO003(ByVal EditMode As EditMode,
                                ByVal Account As DataSet, ByVal aRowIndex As Integer,
                                ByVal VoidData As Boolean, ByVal filterCustId As Boolean,
                                ByRef cmd As DbCommand) As Boolean
        Dim tbSO003C As DataTable = Nothing
        Try
            Dim aServiceIds As String = "-99"
            Dim aProdcutCodes As String = "-99"
            Dim aFaciSeqNo As String = "'X'"
            Dim lst As List(Of DataRow)
            If Not newFlow Then Return True
            If VoidData Then
                lst = New List(Of DataRow)
                For Each rw As DataRow In Account.Tables(FOldProductTableName).Rows
                    lst.Add(rw)
                Next
                'lst.AddRange(Account.Tables(FChangeFaciTableName).Rows)
            Else
                'lst = GetNoChooseFaci(Account)
                lst = GetNoChooseProduct(Account)
            End If
            If lst.Count <= 0 Then
                Return True
            End If
            Dim aPTCode As Integer = -1
            Dim aPTName As String = String.Empty
            Dim aCMCode As Integer = -1
            Dim aCMName As String = String.Empty
            Dim aCustId As String = "-1"
            For Each aRw As DataRow In lst
                aServiceIds = aServiceIds & "," & aRw("ServiceId") & ""
                aCustId = aRw("CUSTID")
            Next
            If filterCustId Then
                tbSO003C = DAO.ExecQry(String.Format(_DAL.GetSO003C, aServiceIds, aCustId))
            Else
                tbSO003C = DAO.ExecQry(String.Format(_DAL.GetNewSO003C,
                                                     aServiceIds, Account.Tables(FDeclaredTableName).Rows(0).Item("SEQNO")))

            End If
            If tbSO003C IsNot Nothing Then
                For Each rw As DataRow In tbSO003C.Rows
                    If Not DBNull.Value.Equals(rw.Item("ProductCode")) Then
                        aProdcutCodes = aProdcutCodes & "," & rw.Item("ProductCode")
                    End If
                    If Not DBNull.Value.Equals(rw.Item("FaciSeqNo")) Then
                        aFaciSeqNo = aFaciSeqNo & ",'" & rw.Item("FaciSeqNo") & "'"
                    End If
                Next
            End If

            'If filterCustId Then
            '    cmd.CommandText = String.Format(_DAL.GetSO003C, aServiceIds, aCustId)
            'Else
            '    cmd.CommandText = String.Format(_DAL.GetNewSO003C, aServiceIds, Account.Tables(FDeclaredTableName).Rows(0).Item("SEQNO"))
            'End If
            'Using dr As DbDataReader = cmd.ExecuteReader()
            '    While dr.Read
            '        If Not DBNull.Value.Equals(dr.Item("ProductCode")) Then
            '            aProdcutCodes = aProdcutCodes & "," & dr.Item("ProductCode")
            '        End If
            '        If Not DBNull.Value.Equals(dr.Item("FaciSeqNo")) Then
            '            aFaciSeqNo = aFaciSeqNo & ",'" & dr.Item("FaciSeqNo") & "'"
            '        End If
            '    End While

            'End Using
            Using tb As DataTable = DAO.ExecQry(_DAL.GetPTCode)
                If tb IsNot Nothing Then
                    aPTCode = tb.Rows(0).Item("CODENO")
                    aPTName = tb.Rows(0).Item("Description")
                    tb.Dispose()
                End If

            End Using

            'cmd.CommandText = _DAL.GetPTCode
            'Using dr As DbDataReader = cmd.ExecuteReader()
            '    dr.Read()
            '    aPTCode = dr.Item("CODENO")
            '    aPTName = dr.Item("Description")
            'End Using
            Using tb As DataTable = DAO.ExecQry(_DAL.GetDefCMCode(LoginInfo))
                If tb IsNot Nothing Then
                    aCMCode = tb.Rows(0).Item("CODENO")
                    aCMName = tb.Rows(0).Item("Description")
                    tb.Dispose()
                End If
            End Using

            'cmd.CommandText = _DAL.GetDefCMCode(LoginInfo, String.Empty)

            'Using dr As DbDataReader = cmd.ExecuteReader

            '    dr.Read()
            '    aCMCode = dr.Item("CODENO")
            '    aCMName = dr.Item("Description")
            'End Using
            If filterCustId Then
                'cmd.CommandText = String.Format(_DAL.ClearSO003,
                '                              aCMCode, aCMName, aPTCode, aPTName, aFaciSeqNo,
                '                              Account.Tables(FNewAccountTableName).Rows(aRowIndex).Item("CustId"),
                '                              aProdcutCodes)
                DAO.ExecNqry(String.Format(_DAL.ClearSO003,
                                              aCMCode, aCMName, aPTCode, aPTName, aFaciSeqNo,
                                              Account.Tables(FNewAccountTableName).Rows(aRowIndex).Item("CustId"),
                                              aProdcutCodes))
            Else
                DAO.ExecNqry(String.Format(_DAL.ClearNewSO003,
                                              aCMCode, aCMName, aPTCode, aPTName,
                                         Account.Tables(FNewAccountTableName).Rows(aRowIndex).Item("UpdEn"),
                                          Account.Tables(FNewAccountTableName).Rows(aRowIndex).Item("UpdTime"),
                                         CType(Account.Tables(FNewAccountTableName).Rows(aRowIndex).Item("NewUpdTime"), Date).ToString("yyyyMMddHHmmss"),
                                             aFaciSeqNo, aProdcutCodes, Account.Tables(FDeclaredTableName).Rows(0).Item("SEQNO")))

                'cmd.CommandText = String.Format(_DAL.ClearNewSO003,
                '                              aCMCode, aCMName, aPTCode, aPTName,
                '                         Account.Tables(FNewAccountTableName).Rows(aRowIndex).Item("UpdEn"),
                '                          Account.Tables(FNewAccountTableName).Rows(aRowIndex).Item("UpdTime"),
                '                         CType(Account.Tables(FNewAccountTableName).Rows(aRowIndex).Item("NewUpdTime"), Date).ToString("yyyyMMddHHmmss"),
                '                             aFaciSeqNo, aProdcutCodes, Account.Tables(FDeclaredTableName).Rows(0).Item("SEQNO"))
            End If

            'cmd.ExecuteNonQuery()

            Return True
        Catch ex As Exception
            Throw
        Finally
            If tbSO003C IsNot Nothing Then
                tbSO003C.Dispose()
                tbSO003C = Nothing
            End If
        End Try
    End Function

    Private Function ClearSO003C(ByVal EditMode As EditMode,
                              ByVal Account As DataSet, ByVal aRowIndex As Integer,
                              ByVal VoidData As Boolean,
                              ByRef cmd As DbCommand) As Boolean
        Try
            Dim aServiceIds As String = "-99"
            Dim lst As List(Of DataRow)
            If VoidData Then
                lst = New List(Of DataRow)
                For Each adr As DataRow In Account.Tables(FOldProductTableName).Rows
                    lst.Add(adr)
                Next
                'lst.AddRange(Account.Tables(FChangeFaciTableName).Rows)
            Else
                'lst = GetNoChooseFaci(Account)
                lst = GetNoChooseProduct(Account)
            End If
            Dim aPTCode As Integer = -1
            Dim aPTName As String = String.Empty
            Dim aCMCode As Integer = -1
            Dim aCMName As String = String.Empty

            For Each aRw As DataRow In lst
                aServiceIds = aServiceIds & "," & aRw("ServiceId")
            Next
            Using tb As DataTable = DAO.ExecQry(_DAL.GetPTCode)
                If tb IsNot Nothing Then
                    aPTCode = tb.Rows(0).Item("CODENO")
                    aPTName = tb.Rows(0).Item("Description")
                End If
                If tb IsNot Nothing Then
                    tb.Dispose()
                End If

            End Using
            'cmd.CommandText = _DAL.GetPTCode
            'Using dr As DbDataReader = cmd.ExecuteReader()
            '    dr.Read()
            '    aPTCode = dr.Item("CODENO")
            '    aPTName = dr.Item("Description")
            'End Using
            Using tb As DataTable = DAO.ExecQry(_DAL.GetDefCMCode(LoginInfo))
                If tb IsNot Nothing Then
                    aCMCode = tb.Rows(0).Item("CODENO")
                    aCMName = tb.Rows(0).Item("Description")
                End If
                If tb IsNot Nothing Then
                    tb.Dispose()
                End If
            End Using

            'cmd.CommandText = _DAL.GetDefCMCode(LoginInfo, String.Empty)

            'Using dr As DbDataReader = cmd.ExecuteReader

            '    dr.Read()
            '    aCMCode = dr.Item("CODENO")
            '    aCMName = dr.Item("Description")
            'End Using
            'cmd.CommandText = String.Format(_DAL.ClearSO003C,
            '                              aCMCode, aCMName, aPTCode, aPTName,
            '                             Account.Tables(FNewAccountTableName).Rows(aRowIndex).Item("UpdEn"),
            '                                Account.Tables(FNewAccountTableName).Rows(aRowIndex).Item("UpdTime"),
            '                            CType(Account.Tables(FNewAccountTableName).Rows(aRowIndex).Item("NewUpdTime"), Date).ToString("yyyyMMddHHmmss"),
            '                              aServiceIds)
            DAO.ExecNqry(String.Format(_DAL.ClearSO003C,
                                          aCMCode, aCMName, aPTCode, aPTName,
                                         Account.Tables(FNewAccountTableName).Rows(aRowIndex).Item("UpdEn"),
                                            Account.Tables(FNewAccountTableName).Rows(aRowIndex).Item("UpdTime"),
                                        CType(Account.Tables(FNewAccountTableName).Rows(aRowIndex).Item("NewUpdTime"), Date).ToString("yyyyMMddHHmmss"),
                                          aServiceIds))
            'DAO.ExecNqry(_DAL.ClearSO003C(aServiceIds,
            '                              CType(Account.Tables(FNewAccountTableName).Rows(aRowIndex).Item("NewUpdTime"), Date).ToString("yyyyMMddHHmmss")),
            '             New Object() {aCMCode, aCMName, aPTCode, aPTName, Account.Tables(FNewAccountTableName).Rows(aRowIndex).Item("UpdEn"),
            '             Account.Tables(FNewAccountTableName).Rows(aRowIndex).Item("UpdTime")})
            'cmd.ExecuteNonQuery()
            Return True
        Catch ex As Exception
            Throw
        End Try
    End Function

    Private Function ClearSO004(ByVal EditMode As EditMode,
                              ByVal Account As DataSet, ByVal aRowIndex As Integer,
                              ByVal VoidData As Boolean,
                              ByRef cmd As DbCommand) As Boolean
        Try
            Dim aSEQNo As String = "'-1X'"
            Dim lst As List(Of DataRow)
            If VoidData Then
                lst = New List(Of DataRow)
                For Each adr As DataRow In Account.Tables(FOldProductTableName).Rows
                    lst.Add(adr)
                Next
                'lst.AddRange(Account.Tables(FChangeFaciTableName).Rows)
            Else
                lst = GetNoChooseFaci(Account)
            End If
            Dim aPTCode As Integer = -1
            Dim aPTName As String = String.Empty
            Dim aCMCode As Integer = -1
            Dim aCMName As String = String.Empty
            aSEQNo = "'-1X'"
            For Each aRw As DataRow In lst
                aSEQNo = aSEQNo & ",'" & aRw("SEQNO") & "'"
            Next
            Using tb As DataTable = DAO.ExecQry(_DAL.GetPTCode)
                If tb IsNot Nothing Then
                    If tb.Rows.Count > 0 Then
                        aPTCode = tb.Rows(0).Item("CODENO")
                        aPTName = tb.Rows(0).Item("Description")
                    End If
                End If
            End Using
            'cmd.CommandText = _DAL.GetPTCode
            'Using dr As DbDataReader = cmd.ExecuteReader()
            '    dr.Read()
            '    aPTCode = dr.Item("CODENO")
            '    aPTName = dr.Item("Description")
            'End Using
            Using tb As DataTable = DAO.ExecQry(_DAL.GetDefCMCode(LoginInfo, String.Empty))
                If tb IsNot Nothing Then
                    If tb.Rows.Count > 0 Then
                        aCMCode = tb.Rows(0).Item("CODENO")
                        aCMName = tb.Rows(0).Item("Description")
                    End If
                End If
            End Using
            'cmd.CommandText = _DAL.GetDefCMCode(LoginInfo, String.Empty)

            'Using dr As DbDataReader = cmd.ExecuteReader

            '    dr.Read()
            '    aCMCode = dr.Item("CODENO")
            '    aCMName = dr.Item("Description")
            'End Using
            'cmd.CommandText = String.Format(_DAL.ClearSO004,
            '                              aCMCode, aCMName, aPTCode, aPTName, aSEQNo)

            'cmd.ExecuteNonQuery(String.Format(_DAL.ClearSO004,
            '                              aCMCode, aCMName, aPTCode, aPTName, aSEQNo))
            DAO.ExecNqry(String.Format(_DAL.ClearSO004,
                                          aCMCode, aCMName, aPTCode, aPTName, aSEQNo))
            Return True
        Catch ex As Exception
            Throw
        End Try

    End Function
    Private Function ClearSO033(ByVal PKBillNoStr As String, aUpdEn As String, aUpdTime As String, aNewUpdTime As Date) As Boolean
        Try
            If String.IsNullOrEmpty(PKBillNoStr) Then
                Return True
            End If
            Dim PtCode As Integer = 1
            Dim ptName As String = Language.ptCash
            Dim UCCode As Integer = 1
            Dim UCName As String = ""
            Using tb As DataTable = DAO.ExecQry(_DAL.GetDefPTCode)
                If tb IsNot Nothing Then
                    PtCode = tb.Rows(0).Item("CODENO")
                    ptName = tb.Rows(0).Item("Description")
                    tb.Dispose()
                End If
            End Using
            'Using dr As DbDataReader = DAO.ExecDtRdr(_DAL.GetDefPTCode)
            '    While dr.Read
            '        PtCode = dr.Item("CODENO")
            '        ptName = dr.Item("Description")
            '    End While
            'End Using
            Using tb As DataTable = DAO.ExecQry(_DAL.GetDefUCCode)
                If tb IsNot Nothing Then
                    UCCode = tb.Rows(0).Item("CODENO")
                    UCName = tb.Rows(0).Item("Description")
                    tb.Dispose()
                End If

            End Using
            'Using dr As DbDataReader = DAO.ExecDtRdr(_DAL.GetDefUCCode)
            '    While dr.Read
            '        UCCode = dr.Item("CODENO")
            '        UCName = dr.Item("Description")
            '    End While
            'End Using
            For Each PKBillNo As String In PKBillNoStr.Replace("'", "").Split(",")
                Dim aServiceType As Object = DAO.ExecSclr(_DAL.GetDataServiceType, New Object() {
                            PKBillNo.Substring(0, 15),
                           Integer.Parse(PKBillNo.Substring(15, PKBillNo.Length - 15))})
                Dim CmCode As String = Nothing
                Dim CmName As String = Nothing
                If aServiceType IsNot Nothing AndAlso Not DBNull.Value.Equals(aServiceType) Then
                    Using tb As DataTable = DAO.ExecQry(_DAL.GetDefCMCode(Me.LoginInfo, aServiceType))
                        If tb IsNot Nothing Then
                            CmCode = tb.Rows(0).Item("CODENO")
                            CmName = tb.Rows(0).Item("Description")
                            tb.Dispose()
                        End If

                    End Using
                    If newFlow Then
                        DAO.ExecNqry(_DAL.ClearSO033,
                                New Object() {Integer.Parse(CmCode), CmName, PtCode, ptName, UCCode, UCName,
                                                   aUpdEn, aUpdTime, aNewUpdTime,
                                              PKBillNo.Substring(0, 15), Integer.Parse(PKBillNo.Substring(15, PKBillNo.Length - 15))})
                    Else
                        DAO.ExecNqry(_DAL.ClearOldSO033,
                                New Object() {Integer.Parse(CmCode), CmName, PtCode, ptName, UCCode, UCName,
                                                   aUpdEn, aUpdTime, aNewUpdTime,
                                              PKBillNo.Substring(0, 15), Integer.Parse(PKBillNo.Substring(15, PKBillNo.Length - 15))})
                    End If

                End If
            Next
        Catch ex As Exception
            Throw
        End Try
        Return True
    End Function
    Private Function UpdSO033(ByVal EditMode As EditMode, ByVal oldAccount As DataTable,
                              ByVal Account As DataSet, ByVal aRowIndex As Integer,
                              ByRef cmd As DbCommand) As Boolean
        Try
            If EditMode = CableSoft.BLL.Utility.EditMode.Edit Then
                If oldAccount IsNot Nothing Then
                    If oldAccount.Rows.Count > 0 Then
                        If Not DBNull.Value.Equals(oldAccount.Rows(0).Item("CitemStr2")) Then
                            ClearSO033(
                              oldAccount.Rows(0).Item("CitemStr2").ToString,
                                   Account.Tables(FNewAccountTableName).Rows(0).Item("UpdEN"),
                                   Account.Tables(FNewAccountTableName).Rows(0).Item("UpdTime"),
                                   Account.Tables(FNewAccountTableName).Rows(0).Item("NewUpdTime"))
                        End If
                    End If
                   
                End If
            End If

            'If Account.Tables(FNewAccountTableName).Rows(0).RowState = DataRowState.Modified Then
            '    If (Not DBNull.Value.Equals(
            '            Account.Tables(FNewAccountTableName).Rows(0).Item("SnactionDate", DataRowVersion.Original))) AndAlso
            '        (Not DBNull.Value.Equals(
            '            Account.Tables(FNewAccountTableName).Rows(0).Item("CitemStr2", DataRowVersion.Original))) Then
            '        If Account.Tables(FNewAccountTableName).Rows(0).Item("CitemStr2", DataRowVersion.Original) <>
            '            Account.Tables(FNewAccountTableName).Rows(0).Item("CitemStr2", DataRowVersion.Current) Then

            '            ClearSO033(
            '                Account.Tables(FNewAccountTableName).Rows(0).Item("CitemStr2", DataRowVersion.Original).ToString,
            '                   Account.Tables(FNewAccountTableName).Rows(0).Item("UpdEN"),
            '                   Account.Tables(FNewAccountTableName).Rows(0).Item("UpdTime"),
            '                   Account.Tables(FNewAccountTableName).Rows(0).Item("NewUpdTime"))

            '        End If
            '    End If
            'End If
            cmd.Parameters.Clear()
            If (Not DBNull.Value.Equals(Account.Tables(FNewAccountTableName).Rows(0).Item("SnactionDate"))) AndAlso
                (Not DBNull.Value.Equals(Account.Tables(FNewAccountTableName).Rows(0).Item("CitemStr2"))) AndAlso
                (Account.Tables(FNewAccountTableName).Rows(0).Item("CitemStr2").ToString.Length > 0) Then
                'If (Not Account.Tables(FNewAccountTableName).Rows(0).HasVersion(DataRowVersion.Original)) OrElse
                '    (DBNull.Value.Equals(Account.Tables(FNewAccountTableName).Rows(0).Item("CitemStr2", DataRowVersion.Original))) OrElse
                '    (Account.Tables(FNewAccountTableName).Rows(0).Item("CitemStr2", DataRowVersion.Original) <>
                '     Account.Tables(FNewAccountTableName).Rows(0).Item("CitemStr2")) Then
                For Each PKBillNo As String In Account.Tables(FNewAccountTableName).Rows(0).Item("CitemStr2").ToString.Replace("'", "").Split(",")
                    If newFlow Then
                        DAO.ExecNqry(_DAL.UpdSO033, New Object() {
                               Account.Tables(FNewAccountTableName).Rows(0).Item("AccountID"),
                               Account.Tables(FNewAccountTableName).Rows(0).Item("BankCode"),
                               Account.Tables(FNewAccountTableName).Rows(0).Item("BankName"),
                               Account.Tables(FNewAccountTableName).Rows(0).Item("CMCode"),
                               Account.Tables(FNewAccountTableName).Rows(0).Item("CMName"),
                               Account.Tables(FNewAccountTableName).Rows(0).Item("PTCode"),
                               Account.Tables(FNewAccountTableName).Rows(0).Item("PTName"),
                               Account.Tables(FNewAccountTableName).Rows(0).Item("UpdEn"),
                               Account.Tables(FNewAccountTableName).Rows(0).Item("UpdTime"),
                               Account.Tables(FNewAccountTableName).Rows(0).Item("NewUpdTime"),
                               PKBillNo.Substring(0, 15), Integer.Parse(PKBillNo.Substring(15, PKBillNo.Length - 15))})
                    Else
                        DAO.ExecNqry(_DAL.UpdOldSO033, New Object() {
                               Account.Tables(FNewAccountTableName).Rows(0).Item("AccountID"),
                               Account.Tables(FNewAccountTableName).Rows(0).Item("BankCode"),
                               Account.Tables(FNewAccountTableName).Rows(0).Item("BankName"),
                               Account.Tables(FNewAccountTableName).Rows(0).Item("CMCode"),
                               Account.Tables(FNewAccountTableName).Rows(0).Item("CMName"),
                               Account.Tables(FNewAccountTableName).Rows(0).Item("PTCode"),
                               Account.Tables(FNewAccountTableName).Rows(0).Item("PTName"),
                               Account.Tables(FNewAccountTableName).Rows(0).Item("UpdEn"),
                               Account.Tables(FNewAccountTableName).Rows(0).Item("UpdTime"),
                               Account.Tables(FNewAccountTableName).Rows(0).Item("NewUpdTime"),
                               SO138_InvSeqNo,
                               PKBillNo.Substring(0, 15), Integer.Parse(PKBillNo.Substring(15, PKBillNo.Length - 15))})
                    End If
                   
                Next
                'End If
            End If

        Catch ex As Exception
            Throw
        End Try
        Return True
    End Function

    Private Function UpdSO003C(ByVal EditMode As EditMode,
                              ByVal Account As DataSet, ByVal aRowIndex As Integer,
                              ByRef cmd As DbCommand) As Boolean
        If chkStopFlag(Account.Tables(FNewAccountTableName).Rows(0)) Then
            Return True
        End If
        If Not newFlow Then Return True
        Try
            Dim a106Rw As DataRow = Account.Tables(FNewAccountTableName).Rows(aRowIndex)
            Dim aServiceIds As String = ""
            'Dim aSeqNos As String = "'X'"
            'Dim aNotSEQNo As String = "'X'"

            For Each aRw As DataRow In Account.Tables(FChangeProductTableName).Rows
                If String.IsNullOrEmpty(aServiceIds) Then
                    aServiceIds = aRw.Item("ServiceId")
                Else
                    aServiceIds = aServiceIds & "," & aRw.Item("ServiceId")
                End If
            Next

            If chkSnactionDate(Account.Tables(FNewAccountTableName).Rows(aRowIndex)) Then
                If aServiceIds = "" Then aServiceIds = "-99"
                'cmd.CommandText = String.Format(_DAL.UpdSO003C,
                '        a106Rw.Item("CMCode"), a106Rw.Item("CMName"),
                '       a106Rw.Item("PTCode"), a106Rw.Item("PTName"),
                '       a106Rw.Item("MasterId"), a106Rw.Item("UpdEn"),
                '       a106Rw.Item("UpdTime"),
                '       CType(a106Rw.Item("NewUpdTime"), Date).ToString("yyyyMMddHHmmss"),
                '       aServiceIds)
                'cmd.ExecuteNonQuery()
                'DAO.ExecNqry(String.Format(_DAL.UpdSO003C,
                '        a106Rw.Item("CMCode"), a106Rw.Item("CMName"),
                '       a106Rw.Item("PTCode"), a106Rw.Item("PTName"),
                '       a106Rw.Item("MasterId"), a106Rw.Item("UpdEn"),
                '       a106Rw.Item("UpdTime"),
                '       CType(a106Rw.Item("NewUpdTime"), Date).ToString("yyyyMMddHHmmss"),
                '       aServiceIds))
                DAO.ExecNqry(_DAL.UpdSO003C(aServiceIds,
                                                          CType(a106Rw.Item("NewUpdTime"), Date).ToString("yyyyMMddHHmmss")),
                                           New Object() {a106Rw.Item("CMCode"), a106Rw.Item("CMName"),
                       a106Rw.Item("PTCode"), a106Rw.Item("PTName"),
                       a106Rw.Item("MasterId"), a106Rw.Item("UpdEn"),
                       a106Rw.Item("UpdTime")})
                ' a106Rw.Item("CMCode"), a106Rw.Item("CMName"),
                'a106Rw.Item("PTCode"), a106Rw.Item("PTName"),
                'a106Rw.Item("MasterId"), a106Rw.Item("UpdEn"),
                'a106Rw.Item("UpdTime"),
                'CType(a106Rw.Item("NewUpdTime"), Date).ToString("yyyyMMddHHmmss"),
                'aServiceIds))
            Else
                'cmd.CommandText = String.Format("UPDATE SO003C SET MASTERID={0} WHERE ServiceId IN ({1}) ",
                '                              a106Rw.Item("MASTERID"), aServiceIds)
                If Not IsACHBank(Account.Tables(FNewAccountTableName).Rows(aRowIndex).Item("BankCode").ToString).ResultBoolean Then
                    If aServiceIds = "" Then aServiceIds = "-99"
                    'cmd.CommandText = _DAL.UpdateSO003C(a106Rw, aServiceIds)
                    'cmd.ExecuteNonQuery()
                    'DAO.ExecNqry(_DAL.UpdateSO003C(a106Rw, aServiceIds))
                    DAO.ExecNqry(_DAL.UpdateSO003C(aServiceIds, CType(a106Rw.Item("NewUpdTime"), Date).ToString("yyyyMMddHHmmss")),
                                 New Object() {a106Rw.Item("MASTERID"), a106Rw.Item("UpdEn"),
                                            a106Rw.Item("UpdTime")})
                Else
                    'Don't Update SO003C if the type is ACH By Kin 2018/02/05
                    If aServiceIds <> "" Then
                        DAO.ExecNqry(_DAL.UpdateACHSO003C, New Object() {aServiceIds, a106Rw.Item("MasterId")})
                    End If
                End If

            End If




            Return True
        Catch ex As Exception
            Throw
        End Try
    End Function

    Private Function UpdSO004(ByVal EditMode As EditMode,
                              ByVal Account As DataSet, ByVal aRowIndex As Integer,
                              ByRef cmd As DbCommand) As Boolean

        If chkStopFlag(Account.Tables(FNewAccountTableName).Rows(0)) Then
            Return True
        End If

        Try
            Dim a106Rw As DataRow = Account.Tables(FNewAccountTableName).Rows(aRowIndex)
            Dim aSEQNo As String = "'-1X'"
            Dim aNotSEQNo As String = "'-1X'"

            For Each aRw As DataRow In Account.Tables(FChangeProductTableName).Rows
                'If (chkSnactionDate(Account.Tables(FCurrectTableName).Rows(aRowIndex))) Then

                'End If
                aSEQNo = aSEQNo & ",'" & aRw.Item("SEQNO") & "'"
            Next

            If chkSnactionDate(Account.Tables(FNewAccountTableName).Rows(aRowIndex)) Then
                'cmd.CommandText = String.Format(_DAL.UpdSO004,
                '                                         a106Rw.Item("AccountId").ToString,
                '       a106Rw.Item("BankCode"), a106Rw.Item("BankName"),
                '       a106Rw.Item("CMCode"), a106Rw.Item("CMName"),
                '       a106Rw.Item("PTCode"), a106Rw.Item("PTName"),
                '       a106Rw.Item("MasterId"), aSEQNo)
                DAO.ExecNqry(String.Format(_DAL.UpdSO004,
                                                         a106Rw.Item("AccountId").ToString,
                       a106Rw.Item("BankCode"), a106Rw.Item("BankName"),
                       a106Rw.Item("CMCode"), a106Rw.Item("CMName"),
                       a106Rw.Item("PTCode"), a106Rw.Item("PTName"),
                       a106Rw.Item("MasterId"), aSEQNo))
            Else
                'cmd.CommandText = String.Format("UPDATE SO004 SET MASTERID={0} WHERE SEQNO IN ({1}) ",
                '                              a106Rw.Item("MASTERID"), aSEQNo)
                'cmd.CommandText = _DAL.UpdateSO004(a106Rw, aSEQNo)
                DAO.ExecNqry(_DAL.UpdateSO004(a106Rw, aSEQNo))
            End If


            'cmd.ExecuteNonQuery()

            Return True
        Catch ex As Exception
            Throw
        End Try
    End Function
    Private Function GetNoChooseProduct(ByVal Account As DataSet) As List(Of DataRow)
        Dim aRet As New List(Of DataRow)
        Try

            For i As Int32 = 0 To Account.Tables(FOldProductTableName).Rows.Count - 1
                Dim ProductCnt = From product In Account.Tables(FChangeProductTableName)
                     Where product.Item("ServiceId") = Account.Tables(FOldProductTableName).Rows(i).Item("ServiceId")
                     Select product.Item("ServiceId")

                If ProductCnt.Count <= 0 Then
                    aRet.Add(Account.Tables(FOldProductTableName).Rows(i))
                End If
            Next


            Return aRet
        Catch ex As Exception
            Throw
        End Try
    End Function
    ''' <summary>
    ''' 取得上次有選此次沒有選擇的設備
    ''' </summary>
    ''' <param name="Account"></param>
    ''' <returns>List(DataRow)</returns>
    ''' <remarks></remarks>
    Private Function GetNoChooseFaci(ByVal Account As DataSet) As List(Of DataRow)
        Dim aRet As New List(Of DataRow)
        Try

            For i As Int32 = 0 To Account.Tables(FOldProductTableName).Rows.Count - 1
                Dim faciCnt = From faci In Account.Tables(FChangeProductTableName)
                     Where faci.Item("SeqNo") = Account.Tables(FOldProductTableName).Rows(i).Item("SeqNo")
                     Select faci.Item("SeqNo")

                If faciCnt.Count <= 0 Then
                    aRet.Add(Account.Tables(FOldProductTableName).Rows(i))
                End If
            Next


            Return aRet
        Catch ex As Exception
            Throw
        End Try

    End Function
    Private Function GetCorrectAccountTable(ByVal EditMode As CableSoft.BLL.Utility.EditMode,
                             ByRef aSourceDs As DataSet) As DataTable
        Try
            Dim aRetAccountTb As DataTable = aSourceDs.Tables(FNewAccountTableName).Copy

            If aRetAccountTb.Rows.Count <= 0 Then Throw New Exception(Language.NoDataUpdate)
            For i As Int32 = 0 To aRetAccountTb.Rows.Count - 1
                If aRetAccountTb.Columns.Contains("CompCode") Then
                    aRetAccountTb.Rows(i).Item("CompCode") = Me.LoginInfo.CompCode
                    aSourceDs.Tables(FNewAccountTableName).Rows(i).Item("CompCode") = Me.LoginInfo.CompCode
                End If
                If aRetAccountTb.Columns.Contains("UpdTime") Then
                    aRetAccountTb.Rows(i).Item("UpdTime") = CableSoft.BLL.Utility.DateTimeUtility.GetDTString(FNowDate)
                    aSourceDs.Tables(FNewAccountTableName).Rows(i).Item("UpdTime") = CableSoft.BLL.Utility.DateTimeUtility.GetDTString(FNowDate)
                    aSourceDs.Tables(FNewAccountTableName).Rows(i).Item("NewUpdTime") = FNowDate
                End If
                If aRetAccountTb.Columns.Contains("UpdEn") Then                    
                    aRetAccountTb.Rows(i).Item("UpdEn") = Me.LoginInfo.EntryName
                    aSourceDs.Tables(FNewAccountTableName).Rows(i).Item("UpdEn") = Me.LoginInfo.EntryName
                End If

                Dim aGetAchCustId As String = GetAchCustId(aRetAccountTb.Rows(i), EditMode)
                If Not String.IsNullOrEmpty(aGetAchCustId) Then
                    aRetAccountTb.Rows(i).Item("AchCustId") = aGetAchCustId
                End If

                If EditMode = CableSoft.BLL.Utility.EditMode.Append Then
                    aRetAccountTb.Rows(i).Item("MasterId") = DAO.ExecSclr(_DAL.TakeSO106SeqNo)
                    aSourceDs.Tables(FNewAccountTableName).Rows(i).Item("MasterId") = aRetAccountTb.Rows(i).Item("MasterId")
                End If

            Next
            Select Case EditMode
                Case EditMode.Append
                    If aRetAccountTb.Columns.Contains("RowId") Then
                        aRetAccountTb.Columns.Remove(aRetAccountTb.Columns("RowId"))
                    End If
                    If aRetAccountTb.Columns.Contains("CTID") Then
                        aRetAccountTb.Columns.Remove(aRetAccountTb.Columns("CTID"))
                    End If
                Case EditMode.Edit
                    If aRetAccountTb.Columns.Contains("RowId") Then
                        aRetAccountTb.Columns.Remove(aRetAccountTb.Columns("RowId"))
                    End If
                    If aRetAccountTb.Columns.Contains("CTID") Then
                        aRetAccountTb.Columns.Remove(aRetAccountTb.Columns("CTID"))
                    End If
                    If aRetAccountTb.Columns.Contains("MasterId") Then
                        aRetAccountTb.Columns.Remove(aRetAccountTb.Columns("MasterId"))
                    End If
                Case EditMode.Delete
                    'If Not aRetTb.Columns.Contains("RowId") Then Throw New Exception("無傳入RowId")
            End Select


            Return aRetAccountTb
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Public Function IsACHBank(ByVal aBankCode As String, ByVal blnStartPost As Boolean) As RIAResult
        Try
            'Dim cnt As Int32 = Int32.Parse(DAO.ExecSclr(String.Format("Select Count(*) From CD018 Where CodeNo ={0}  And PRGNAME LIKE 'ACH%'",
            '                                                 aBankCode)))


            Dim cnt As Int32 = Int32.Parse(DAO.ExecSclr(_DAL.IsACHBank(blnStartPost), New Object() {Integer.Parse(aBankCode)}))
            If cnt <= 0 Then
                Return New RIAResult() With {.ErrorCode = 0, .ErrorMessage = String.Empty, .ResultBoolean = False}
            End If
            Return New RIAResult() With {.ErrorCode = 0, .ErrorMessage = String.Empty, .ResultBoolean = True}
        Catch ex As Exception
            Return New RIAResult() With {.ErrorCode = -1, .ErrorMessage = ex.ToString, .ResultBoolean = False}
        End Try
    End Function
    Public Function IsACHBank(ByVal aBankCode As String) As RIAResult
        Try
            'Dim cnt As Int32 = Int32.Parse(DAO.ExecSclr(String.Format("Select Count(*) From CD018 Where CodeNo ={0}  And PRGNAME LIKE 'ACH%'",
            '                                                 aBankCode)))
            ' Dim startPost As Boolean = Integer.Parse(DAO.ExecSclr("Select Nvl(StartPost,0) From SO041 Where SysID = " & LoginInfo.CompCode)) = 1
            Dim startPost As Boolean = Integer.Parse(DAO.ExecSclr(_DAL.QueryStartPost, New Object() {LoginInfo.CompCode})) = 1
            Dim cnt As Int32 = Int32.Parse(DAO.ExecSclr(_DAL.IsACHBank(startPost), New Object() {Integer.Parse(aBankCode)}))
            If cnt <= 0 Then
                Return New RIAResult() With {.ErrorCode = 0, .ErrorMessage = String.Empty, .ResultBoolean = False}
            End If
            Return New RIAResult() With {.ErrorCode = 0, .ErrorMessage = String.Empty, .ResultBoolean = True}
        Catch ex As Exception
            Return New RIAResult() With {.ErrorCode = -1, .ErrorMessage = ex.ToString, .ResultBoolean = False}
        End Try
    End Function
    Private Function GetAchCustId(ByVal rwAccount As DataRow, ByVal EditMode As EditMode) As String
        Dim aACHCustID As Boolean = False
        Try
            'Dim aAchCnt As Int32 = Int32.Parse(DAO.ExecSclr(String.Format("Select Count(*) From CD018 Where CodeNo ={0}  And PRGNAME LIKE 'ACH%'",
            '                                                  rwAccount.Item("BankCode"))))
            'Dim startPost As Boolean = Integer.Parse(DAO.ExecSclr("Select Nvl(StartPost,0) From SO041 Where SysID = " & LoginInfo.CompCode)) = 1
            Dim startPost As Boolean = Integer.Parse(DAO.ExecSclr(_DAL.QueryStartPost, New Object() {LoginInfo.CompCode})) = 1
            Dim aAchCnt As Int32 = Int32.Parse(DAO.ExecSclr(_DAL.IsACHBank(startPost), New Object() {rwAccount.Item("BankCode")}))
            If aAchCnt <= 0 Then
                Return Nothing
            End If
            aACHCustID = Int32.Parse(DAO.ExecSclr(_DAL.GetACHCustId,
                                                           New Object() {Me.LoginInfo.CompCode})) = 1

            If (Not rwAccount.IsNull("ACHCustId")) AndAlso (Not String.IsNullOrEmpty(rwAccount.Item("ACHCustId"))) AndAlso
                (EditMode = CableSoft.BLL.Utility.EditMode.Append) Then
                Return Nothing
            Else
                If (EditMode = CableSoft.BLL.Utility.EditMode.Edit) AndAlso
                    (DBNull.Value.Equals(rwAccount.Item("SendDate"))) Then
                    If rwAccount.HasVersion(DataRowVersion.Original) Then
                        If rwAccount.Item("AccountID") = rwAccount.Item("AccountID", DataRowVersion.Original) Then
                            Return Nothing
                        End If

                    End If
                Else
                    If Not DBNull.Value.Equals(rwAccount.Item("SendDate")) Then
                        Return Nothing
                    End If
                End If

                Dim aAchTotal As String = String.Empty
                Dim aMasterId As String = "-1"
                Dim aACHHeadCode As String = GetMaxAchNo(rwAccount.Item("ACHTNO").ToString,
                                                         aACHCustID)
                'If rwAccount.IsNull("ACHTNO") OrElse String.IsNullOrEmpty(rwAccount.Item("ACHTNO")) Then
                '    aACHHeadCode = String.Empty
                'Else
                '    aACHHeadCode = rwAccount("ACHTNO").ToString.Replace("'", "").Substring(0, 3)
                'End If
                'Dim aACHHeadCode As String = aRw("ACHTNO").ToString.Replace("'", "").Substring(0, 3)

                'Dim objAchHead As Object = DAO.ExecSclr("select ACHHeadCode from SO041")
                'If objAchHead IsNot Nothing Then
                '    aACHHeadCode = objAchHead.ToString
                'End If
                'If String.IsNullOrEmpty(aACHHeadCode) Then
                '    Return Nothing
                'End If

                If EditMode = CableSoft.BLL.Utility.EditMode.Edit Then
                    aMasterId = rwAccount.Item("MasterId").ToString
                End If



                If Not aACHCustID Then
                    Return aACHHeadCode & Right(New String("0", 8) & rwAccount.Item("CUSTID"), 8)
                End If
                'aAchTotal = DAO.ExecSclr(String.Format("Select Count(*) + 1 From SO106 " &
                '                       " Where SUBSTR(LPAD(AccountId,30,'0'),25,6) = '{0}' " & _
                '                       " And MasterId <> {1}" & _
                '                       " AND SUBSTR(ACHCUSTID,4,6) = '{2}' ",
                '                       Right(rwAccount.Item("AccountId").ToString, 6),
                '                       aMasterId,
                '                       Right(rwAccount.Item("AccountId").ToString, 6)))
                Dim aWhereIn As String = Nothing
                For i As Int32 = 1 To 99
                    If String.IsNullOrEmpty(aWhereIn) Then
                        aWhereIn = String.Format("'{0}'", aACHHeadCode &
                                               Right(rwAccount.Item("AccountId").ToString.PadLeft(6, "0"c).ToString, 6) &
                                                i.ToString.PadLeft(2, "0"c))
                    Else
                        aWhereIn = String.Format("{0},{1}", aWhereIn,
                                                 String.Format("'{0}'", aACHHeadCode &
                                               Right(rwAccount.Item("AccountId").ToString.PadLeft(6, "0"c).ToString, 6) &
                                                i.ToString.PadLeft(2, "0"c)))

                    End If
                Next
                '5354 檢核01-99那個流水號沒用過 By Kin 2012/09/05
                'Dim aSQL As String = String.Format("SELECT ACHCUSTID From SO106 " &
                '                       " Where SUBSTR(LPAD(AccountId,30,'0'),25,6) = '{0}' " & _
                '                       " And MasterId <> {1}" & _
                '                       " AND SUBSTR(ACHCUSTID,4,6) = '{2}' " & _
                '                       " AND ACHCUSTID IN ({3}) ORDER BY ACHCUSTID ",
                '                       Right(rwAccount.Item("AccountId").ToString, 6),
                '                       aMasterId,
                '                       Right(rwAccount.Item("AccountId").ToString, 6), aWhereIn)
                Dim aSQL As String = _DAL.QueryNoUseAchCustId(Right(rwAccount.Item("AccountId").ToString, 6),
                                                                     aMasterId, aWhereIn)

                Using dtUseACH As DataTable = DAO.ExecQry(aSQL)
                    aAchTotal = "01"
                    If dtUseACH.Rows.Count > 0 Then
                        Dim aFind As Boolean = False
                        For i As Int32 = 1 To 99
                            If aFind = True Then
                                Exit For
                            End If
                            aAchTotal = i.ToString.PadLeft(2, "0")
                            For Each rw As DataRow In dtUseACH.Rows
                                aFind = True
                                If Right(rw.Item("ACHCUSTID"), 2) = aAchTotal Then
                                    aFind = False
                                    Exit For
                                End If
                            Next
                        Next
                    End If
                End Using
                Return aACHHeadCode & Right(rwAccount.Item("AccountId").ToString.PadLeft(6, "0"c).ToString, 6) &
                    Right(aAchTotal.PadLeft(2, "0"c), 2)
            End If

        Catch ex As Exception
            Throw ex
        End Try

    End Function
    Private Function GetMaxAchNo(ByVal strMax As String, ByVal blnAchCustId As Boolean) As String
        Dim aryStr As String()
        Dim strAchMax As String

        Dim aACHHeadCode As String = Nothing
        Try
            If blnAchCustId Then
                aACHHeadCode = DAO.ExecSclr(_DAL.GetAchHeadCode,
                                            New Object() {Me.LoginInfo.CompCode}).ToString
                If Not String.IsNullOrEmpty(aACHHeadCode) Then
                    Return aACHHeadCode
                End If
            End If
            aryStr = strMax.Replace("'", "").Split(",")
            strAchMax = aryStr(0)
            For i As Int32 = 0 To aryStr.Count - 1
                If Int32.Parse(aryStr(i)) > (strAchMax) Then
                    strAchMax = aryStr(i)
                End If
            Next
        Catch ex As Exception
            Throw ex
        End Try


        Return strAchMax
    End Function
    Private Function HavePK(ByVal aEditMode As EditMode, ByVal aTB As DataTable) As Boolean
        Try
            If aEditMode = EditMode.Append Then
                Return True
            End If
            If aTB.Columns.Contains("ROWID") Then
                Return True
            End If
            If aTB.Columns.Contains("CTID") Then
                Return True
            End If
            For Each s As String In FPKField.Split(",")
                If Not aTB.Columns.Contains(s) Then
                    Return False
                End If
            Next
            Return True
        Catch ex As Exception
            Throw
        End Try
    End Function
    Public Function ChkCMDataOk(ByVal Account As DataSet) As RIAResult
        Try
            Return ChkMustCMData(Account.Tables(FNewAccountTableName).Rows(0))
            'If ChkMustCMData(Account.Tables(FNewAccountTableName).Rows(0)) Then
            '    Return New RIAResult() With {.ErrorCode = 0, .ErrorMessage = String.Empty, .ResultBoolean = True}
            'Else
            '    Return New RIAResult() With {.ErrorCode = -1, .ErrorMessage = Language.CMDataError, .ResultBoolean = False}
            'End If
        Catch ex As Exception
            Return New RIAResult() With {.ErrorCode = -1, .ErrorMessage = ex.Message, .ResultBoolean = False}
        End Try
    End Function

    Private Function UpdSO106(ByVal aAccount As DataSet, ByVal aEditMode As EditMode) As Boolean
        Try
            Select Case aEditMode
                Case EditMode.Edit
                    Dim aWhere As String = String.Empty
                    Dim aNow As String = CableSoft.BLL.Utility.DateTimeUtility.GetDTString(Date.Now)
                    For i As Int32 = 0 To aAccount.Tables(FNewAccountTableName).Rows.Count - 1
                        aAccount.Tables(FNewAccountTableName).Rows(i).Item("UpdTime") = aNow
                    Next

            End Select

            Return True
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Private Function HaveMustField(ByVal aDT As DataTable) As Boolean
        Try
            If Not aDT.Columns.Contains("Proposer") Then
                Throw New Exception(Language.NoProposer)
            Else
                If aDT.Rows(0).IsNull("Proposer") Then Throw New Exception(Language.NoProposerData)
            End If
            If Not aDT.Columns.Contains("PropDate") Then
                Throw New Exception(Language.NoPropDate)
            Else
                If aDT.Rows(0).IsNull("PropDate") Then Throw New Exception(Language.NoPropDateData)
            End If
            If Not aDT.Columns.Contains("CMCode") Then
                Throw New Exception(Language.NoCMCodeField)
            Else
                If aDT.Rows(0).IsNull("CMCode") Then
                    Throw New Exception(Language.NoCMCodeData)
                End If
            End If
            If Not aDT.Columns.Contains("PTCode") Then
                Throw New Exception(Language.NoPTCodeField)
            Else
                If aDT.Rows(0).IsNull("PTCode") Then
                    Throw New Exception(Language.NoPtCodeData)
                End If
            End If
            If Not aDT.Columns.Contains("BankCode") Then
                Throw New Exception("無銀行欄位！")
            Else
                If aDT.Rows(0).IsNull("BankCode") Then
                    Throw New Exception("無銀行資料！")
                End If
            End If
            If Not aDT.Columns.Contains("AccountID") Then
                Throw New Exception("無帳號欄位！")
            Else
                If aDT.Rows(0).IsNull("AccountID") Then
                    Throw New Exception("無帳號資料！")
                End If
            End If
            If aDT.Columns.Contains("StopFlag") Then
                If (Not aDT.Rows(0).IsNull("StopFlag")) AndAlso (Int32.Parse(aDT.Rows(0).Item("StopFlag"))) = 1 Then
                    If Not aDT.Columns.Contains("StopDate") Then
                        Throw New Exception("無停用日期欄位！")
                    Else
                        If aDT.Rows(0).IsNull("StopDate") Then
                            Throw New Exception("無停用日期資料！")
                        End If
                    End If
                End If
            End If
            Return True
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Private Function ChkMustCMData(ByVal aRow As DataRow) As RIAResult
        Dim result As New RIAResult With {.ResultBoolean = False, .ErrorCode = -1, .ErrorMessage = Nothing}
        Try
            Dim aCMRefNo As Int32 = Int32.Parse("0" & DAO.ExecSclr(_DAL.GetCMRefNo, aRow.Item("CMCode")))
            Select Case aCMRefNo
                Case 2
                    If Not aRow.Table.Columns.Contains("AccountName") Then
                        result.ErrorMessage = Language.noAccountNameField
                        Return result

                    Else
                        If (aRow.IsNull("AccountName")) OrElse (String.IsNullOrEmpty(aRow.Item("AccountName"))) Then
                            result.ErrorMessage = Language.noAccountNameData
                            Return result
                        End If
                    End If
                    If Not aRow.Table.Columns.Contains("AccountNameID") Then
                        result.ErrorMessage = Language.noAccountNameIDField
                        Return result
                    Else
                        If (aRow.IsNull("AccountNameID")) OrElse (String.IsNullOrEmpty(aRow.Item("AccountNameID"))) Then
                            result.ErrorMessage = Language.noAccountNameIDData
                            Return result
                        End If
                    End If
                    If (aRow.Table.Columns.Contains("CardCode")) AndAlso
                        (Not aRow.IsNull("CardCode")) Then
                        Dim aCardNoLen As Int32 = Int32.Parse("0" & DAO.ExecSclr(_DAL.GetCardNoLen,
                                                               aRow.Item("CardCode")))
                        If Convert.ToString(aRow.Item("AccountID")).Length <> aCardNoLen Then
                            result.ErrorMessage = String.Format(Language.AccountIdLimit, aCardNoLen)
                            Return result
                        End If
                    Else
                        If Not aRow.Table.Columns.Contains("BankCode") Then
                            result.ErrorMessage = Language.noBankCodeField
                            Return result
                        Else
                            If (aRow.IsNull("BankCode")) OrElse (String.IsNullOrEmpty(aRow.Item("BankCode"))) Then
                                result.ErrorMessage = Language.noBankCodeData
                                Return result
                            Else
                                Dim aActLength As Int32 = Int32.Parse(DAO.ExecSclr(_DAL.GetActLength, aRow.Item("BankCode")))
                                If Convert.ToString(aRow.Item("AccountID")).Length <> aActLength Then
                                    result.ErrorMessage = String.Format(Language.AccountIdLimit, aActLength)
                                    result.ResultBoolean = False
                                    Return result
                                End If
                            End If
                        End If
                    End If
                Case 4
                    'D.	信用卡別(CardCode,CardName),信用卡有效期限(StopYM),帳號所有人(AccountName),帳號所有人ID(AccountNameID)
                    If (Not aRow.Table.Columns.Contains("CardCode")) OrElse
                        (Not aRow.Table.Columns.Contains("CardName")) Then
                        result.ErrorMessage = Language.noCardCodeField
                        Return result

                    Else
                        If (aRow.IsNull("CardCode")) OrElse (aRow.IsNull("CardName")) Then
                            result.ErrorMessage = Language.noCardCodeData
                            Return result
                        End If
                    End If
                    If (Not aRow.Table.Columns.Contains("StopYM")) Then
                        result.ErrorMessage = Language.noStopYMField
                        Return result
                    Else
                        If (aRow.IsNull("StopYM")) OrElse (String.IsNullOrEmpty(aRow.Item("StopYM"))) Then
                            result.ErrorMessage = Language.noStopYMData
                            Return result
                        End If
                    End If
                    If Not aRow.Table.Columns.Contains("AccountName") Then
                        result.ErrorMessage = Language.noAccountNameField
                        Return result
                    Else
                        If (aRow.IsNull("AccountName")) OrElse (String.IsNullOrEmpty(aRow.Item("AccountName"))) Then
                            result.ErrorMessage = Language.noAccountNameData
                            Return result
                        End If
                    End If
                    If Not aRow.Table.Columns.Contains("AccountNameID") Then
                        result.ErrorMessage = Language.noAccountNameIDField
                        Return result
                    Else
                        If (aRow.IsNull("AccountNameID")) OrElse (String.IsNullOrEmpty(aRow.Item("AccountNameID"))) Then
                            result.ErrorMessage = Language.noAccountNameIDData
                            Return result
                        End If
                    End If

            End Select
            result.ResultBoolean = True
            result.ErrorCode = 0
            Return result
        Catch ex As Exception
            Throw
        End Try


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
                ' TODO: 處置 Managed 狀態 (Managed 物件)。
                If Language IsNot Nothing Then
                    Language.Dispose()
                    Language = Nothing
                End If
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
Friend Structure AchStatus
    Public Property ACHTNo As String
    Public Property ACHTDesc As String
    Public Property UpdateType As AchUpdateType
    Public Property CitemCode As Object
    Public Property CitemName As Object
End Structure
Friend Enum AchUpdateType
    AddAuthorize = 0
    CancelAuthorize = 1
    ChangeCitem = 2
End Enum