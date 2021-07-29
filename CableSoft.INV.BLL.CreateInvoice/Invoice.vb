Imports System.Data.Common
Imports CableSoft.BLL.BillingAPI
Imports CableSoft.BLL.Utility
Public Class Invoice
    Inherits BLLBasic
    Implements IDisposable

    Private _DAL As New InvoiceDALMultiDB(Me.LoginInfo.Provider)
    Private Language As New CableSoft.BLL.Language.SO61.Invoice
    Private executeTime As Date = Date.Now
    Private pNewFlow As Integer = 0
    Private misCompCode As Integer = 0
    Public Sub New()
    End Sub
    Public Sub New(ByVal LoginInfo As CableSoft.BLL.Utility.LoginInfo)
        MyBase.New(LoginInfo)
        ReadInv001()
    End Sub
    Public Sub New(ByVal LoginInfo As CableSoft.BLL.Utility.LoginInfo, ByVal DAO As CableSoft.Utility.DataAccess.DAO)
        MyBase.New(LoginInfo, DAO)
        ReadInv001()
    End Sub
    Public Sub New(ByVal LoginInfo As CableSoft.BLL.Utility.LoginInfo, ByVal DBConnection As System.Data.Common.DbConnection)
        MyBase.New(LoginInfo, DBConnection)
        ReadInv001()
    End Sub
    Private Sub ReadInv001()
        Using dt As DataTable = DAO.ExecQry(_DAL.QueryInv001, New Object() {LoginInfo.CompCode.ToString()})
            misCompCode = LoginInfo.CompCode
            If Not DBNull.Value.Equals(dt.Rows(0).Item("NewFlow")) Then
                pNewFlow = Integer.Parse(dt.Rows(0).Item("NewFlow"))
            End If
            If Not DBNull.Value.Equals(dt.Rows(0).Item("misowner")) Then
                _DAL.misOwner = Replace(dt.Rows(0).Item("misowner").ToString, ".", "") & "."
            End If
            If Not DBNull.Value.Equals(dt.Rows(0).Item("sysid")) Then
                misCompCode = Integer.Parse(dt.Rows(0).Item("sysid"))
            End If
        End Using
    End Sub
    Public Function QueryEditData() As DataSet
        Dim ds As New DataSet
        Dim aServiceTypestr As String = Nothing
        Try
            Using tb001 As DataTable = DAO.ExecQry(_DAL.QueryInv001, New Object() {LoginInfo.CompCode.ToString})
                If tb001.Rows.Count = 1 Then
                    If Not DBNull.Value.Equals(tb001.Rows(0).Item("ServiceTypeStr")) Then
                        For Each o As String In tb001.Rows(0).Item("ServiceTypeStr").ToString.Split(",")
                            If String.IsNullOrEmpty(aServiceTypestr) Then
                                aServiceTypestr = "'" & o.Replace("'", "") & "'"
                            Else
                                aServiceTypestr = aServiceTypestr & ",'" & o.Replace("'", "") & "'"
                            End If
                        Next
                    End If
                End If
            End Using

            Dim tbInv005 As DataTable = DAO.ExecQry(_DAL.QueryINV005, New Object() {LoginInfo.CompCode.ToString}).Copy
            tbInv005.TableName = "INV005"
            Dim tbServiceType As DataTable = DAO.ExecQry(_DAL.QueryINV001ServiceType(aServiceTypestr)).Copy
            tbServiceType.TableName = "SERVICETYPE"
            ds.Tables.Add(tbInv005)
            ds.Tables.Add(tbServiceType)

            Return ds.Copy
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Public Function CanEdit(ByVal InvId As String) As RIAResult
        Dim result As New RIAResult With {.ErrorCode = -1, .ErrorMessage = "initial Error", .ResultBoolean = False}

        Try
            Using tbINV007 As DataTable = DAO.ExecQry(_DAL.QueryINV007,
                                                  New Object() {InvId})
                If tbINV007.Rows.Count = 0 Then
                    result.ErrorMessage = Language.NoInvModify
                    Return result
                End If
                If tbINV007.Rows(0).Item("IsObsolete").ToString = "Y" Then
                    result.ErrorMessage = Language.ObsoleteYes
                    Return result
                End If
                If tbINV007.Rows(0).Item("CANMODIFY").ToString <> "Y" Then
                    result.ErrorMessage = Language.NOMODIFY
                    Return result
                End If
                If Integer.Parse(DAO.ExecNqry(_DAL.QueryIsLock, New Object() {LoginInfo.CompCode.ToString(),
                             tbINV007.Rows(0).Item("INVID")}) & "0") > 0 Then
                    result.ErrorMessage = Language.INVDue
                    Return result
                End If

            End Using
            result.ResultBoolean = True
            result.ErrorMessage = Nothing
            result.ErrorCode = 0
        Catch ex As Exception
            Throw ex
        End Try

        Return result
    End Function
    Public Function QueryAllData(ByVal editMode As EditMode, ByVal Invid As String) As DataSet

        Dim ds As New DataSet()
        If editMode = EditMode.Append Then
            Invid = "X9"
        End If
        Try

            Dim tbInv007 As DataTable = DAO.ExecQry(_DAL.QueryINV007, New Object() {Invid}).Copy
            Dim tbInv008 As DataTable = DAO.ExecQry(_DAL.QueryINV008, New Object() {Invid}).Copy
            Dim tbInv008A As DataTable = DAO.ExecQry(_DAL.QueryINV008A, New Object() {Invid}).Copy
            Dim tbInv028 As DataTable = DAO.ExecQry(_DAL.QueryINV028, New Object() {LoginInfo.CompCode.ToString()}).Copy
            Dim tbInv041 As DataTable = DAO.ExecQry(_DAL.QueryINV041, New Object() {LoginInfo.CompCode.ToString()}).Copy
            Dim tbInv001 As DataTable = DAO.ExecQry(_DAL.QueryInv001, New Object() {LoginInfo.CompCode.ToString()}).Copy
            tbInv007.TableName = "INV007"
            tbInv008.TableName = "INV008"
            tbInv008A.TableName = "INV008A"
            tbInv028.TableName = "INV028"
            tbInv041.TableName = "INV041"
            tbInv001.TableName = "INV001"
            ds.Tables.Add(tbInv007)
            ds.Tables.Add(tbInv008)
            ds.Tables.Add(tbInv008A)
            ds.Tables.Add(tbInv028)
            ds.Tables.Add(tbInv041)
            ds.Tables.Add(tbInv001)
        Catch ex As Exception
            Throw ex
        End Try
        Return ds.Copy
    End Function
    Public Function QueryINV099(ByVal InvDate As Date) As DataSet
        Dim ds As New DataSet()
        Try
            Dim mm As String = InvDate.ToString("MM")
            Dim yyyy As String = InvDate.ToString("yyyy")


            If (Integer.Parse(InvDate.ToString("MM")) Mod 2 = 0) Then
                mm = Right("00" & Integer.Parse(mm) - 1, 2)
            End If

            Dim tbInv099 As DataTable = DAO.ExecQry(_DAL.QueryINV099, New Object() {
               LoginInfo.CompCode.ToString(), yyyy & mm, InvDate.ToString("yyyyMMdd")}).Copy
            ds.Tables.Add(tbInv099)
        Catch ex As Exception
            Throw ex
        End Try


        Return ds.Copy

    End Function
    Public Function QuerySOCustId(ByVal custid As String) As DataSet
        Dim ds As New DataSet
        Try
            Dim tbSO001 As DataTable = DAO.ExecQry(_DAL.QuerySO001(custid), New Object() {
                  misCompCode}).Copy

            tbSO001.TableName = "SO001"
            ds.Tables.Add(tbSO001)
        Catch ex As Exception
            Throw ex

        End Try
        Return ds
    End Function
    Public Function QuerySOBill(ByVal existsBillNo As String, ByVal custid As String, ByVal invseqno As String, ByVal guino As String, ByVal newflow As Integer) As DataSet
        Dim invServiceTypestr As String = "'x'"
        Dim ds As New DataSet

        Dim tbBill As DataTable = Nothing
        If String.IsNullOrEmpty(existsBillNo) Then existsBillNo = "'x1'"
        Try
            Using tb As DataTable = DAO.ExecQry(_DAL.QueryInv001, New Object() {LoginInfo.CompCode.ToString()})
                For Each rw As DataRow In tb.Rows
                    If Not IsDBNull(rw.Item("ServiceTypeStr")) Then
                        For Each o As String In rw.Item("ServiceTypeStr").ToString.Split(",")
                            invServiceTypestr = String.Format("{0},'{1}'", invServiceTypestr, o)
                        Next
                    End If
                Next
                tb.Dispose()
            End Using
            Select Case newflow
                Case 1
                    tbBill = DAO.ExecQry(_DAL.QuerySOBill(invServiceTypestr, existsBillNo),
                                                  New Object() {Integer.Parse(custid), Integer.Parse(invseqno), guino, LoginInfo.CompCode.ToString}).Copy
                Case Else
                    tbBill = DAO.ExecQry(_DAL.QueryOldSOBill(invServiceTypestr, existsBillNo),
                                                  New Object() {Integer.Parse(custid), guino, LoginInfo.CompCode.ToString}).Copy
            End Select

            tbBill.TableName = "BILL"
            tbBill.Columns.Add("CHOOSE", GetType(Boolean))
            For i As Integer = 0 To tbBill.Rows.Count - 1
                tbBill.Rows(i).Item("CHOOSE") = False
            Next
            tbBill.AcceptChanges()
            ds.Tables.Add(tbBill)
        Catch ex As Exception
            Throw ex
        End Try

        Return ds.Copy
    End Function

    Public Function QuerySOBill(ByVal existsBillNo As String, ByVal custid As String, ByVal invseqno As String, ByVal guino As String) As DataSet
        Return QuerySOBill(existsBillNo, custid, invseqno, guino, pNewFlow)
    End Function
    Private Function getYearMonth(ByVal InvDate As Date) As String
        Dim YearMonth As String = Date.Parse(InvDate).ToString("MM")
        If (Integer.Parse(YearMonth) Mod 2 = 0) Then
            YearMonth = Date.Parse(InvDate).ToString("yyyy") & Right("00" & (Integer.Parse(YearMonth) - 1), 2)
        Else
            YearMonth = Date.Parse(InvDate).ToString("yyyyMM")
        End If
        Return YearMonth
    End Function
    Private Function checkCanCreate(ByVal tbINV007 As DataTable, ByVal tbINV008 As DataTable, ByVal aAutoCreateNum As Integer) As RIAResult
        Dim result As New RIAResult With {.ResultBoolean = False, .ErrorCode = -1, .ErrorMessage = "checkCanCreate"}
        Try
            'Dim YearMonth As String = Date.Parse(tbINV007.Rows(0).Item("INVDATE")).ToString("MM")
            'If (Integer.Parse(YearMonth) Mod 2 = 0) Then
            '    YearMonth = Date.Parse(tbINV007.Rows(0).Item("INVDATE")).ToString("yyyy") & Right("00" & (Integer.Parse(YearMonth) - 1), 2)
            'Else
            '    YearMonth = Date.Parse(tbINV007.Rows(0).Item("INVDATE")).ToString("yyyyMM")
            'End If
            Dim YearMonth As String = getYearMonth(tbINV007.Rows(0).Item("INVDATE"))
            Dim PREFIX As String = tbINV007.Rows(0).Item("INVID").ToString.Substring(0, 2)
            Using tbCurrentInv099 As DataTable = DAO.ExecQry(_DAL.QueryCurrentInv099, New Object() {
                                       LoginInfo.CompCode.ToString,
                                       Right(tbINV007.Rows(0).Item("INVID").ToString, 8),
                                       YearMonth, PREFIX
                                  })


                If tbCurrentInv099.Rows.Count = 0 Then
                    result.ErrorMessage = Language.NOINV099
                    Return result
                End If
                Dim STARTNUM As String = tbCurrentInv099.Rows(0).Item("STARTNUM").ToString

                Dim aAvalibleCount As Integer = DAO.ExecSclr(_DAL.GetCurrentAvailableInvCount, New Object() {
                                                    LoginInfo.CompCode.ToString, YearMonth, PREFIX, STARTNUM
                                                })
                Dim readyCreateNum As Integer = Math.Ceiling(tbINV008.Rows.Count / aAutoCreateNum)
                If aAvalibleCount < readyCreateNum Then
                    result.ErrorMessage = String.Format(Language.INV099Waring, readyCreateNum, aAvalibleCount)

                    Return result
                End If

            End Using



            result.ResultBoolean = True

        Catch ex As Exception
            Throw
        End Try
        Return result
    End Function
    Private Sub DisplaceScrInv007(ByVal tbScrInv007 As DataTable, ByVal tbDBInv007 As DataTable)
        Try
            If tbDBInv007.Rows.Count > 0 Then
                tbScrInv007.Rows(0).Item("IsPast") = tbDBInv007.Rows(0).Item("IsPast")
                tbScrInv007.Rows(0).Item("ObsoleteId") = tbDBInv007.Rows(0).Item("ObsoleteId")
                tbScrInv007.Rows(0).Item("ObsoleteReason") = tbDBInv007.Rows(0).Item("ObsoleteReason")
                tbScrInv007.Rows(0).Item("PrintFun") = tbDBInv007.Rows(0).Item("PrintFun")
                tbScrInv007.Rows(0).Item("PrintTime") = tbDBInv007.Rows(0).Item("PrintTime")
                tbScrInv007.Rows(0).Item("PrizeType") = tbDBInv007.Rows(0).Item("PrizeType")
                tbScrInv007.Rows(0).Item("UploadFlag") = tbDBInv007.Rows(0).Item("UploadFlag")
                tbScrInv007.Rows(0).Item("PrintEn") = tbDBInv007.Rows(0).Item("PrintEn")
                tbScrInv007.Rows(0).Item("UploadTime") = tbDBInv007.Rows(0).Item("UploadTime")
                tbScrInv007.Rows(0).Item("ObUploadFlag") = tbDBInv007.Rows(0).Item("ObUploadFlag")
                tbScrInv007.Rows(0).Item("ObUploadTime") = tbDBInv007.Rows(0).Item("ObUploadTime")
                tbScrInv007.Rows(0).Item("CreateInvDate") = tbDBInv007.Rows(0).Item("CreateInvDate")
                tbScrInv007.Rows(0).Item("CarrierType") = tbDBInv007.Rows(0).Item("CarrierType")
                tbScrInv007.Rows(0).Item("CarrierId1") = tbDBInv007.Rows(0).Item("CarrierId1")
                tbScrInv007.Rows(0).Item("CarrierId2") = tbDBInv007.Rows(0).Item("CarrierId2")
                tbScrInv007.Rows(0).Item("RandomNum") = tbDBInv007.Rows(0).Item("RandomNum")
                tbScrInv007.Rows(0).Item("A_CarrierId1") = tbDBInv007.Rows(0).Item("A_CarrierId1")
                tbScrInv007.Rows(0).Item("A_CarrierId2") = tbDBInv007.Rows(0).Item("A_CarrierId2")
                tbScrInv007.Rows(0).Item("DestroyFlag") = tbDBInv007.Rows(0).Item("DestroyFlag")
                tbScrInv007.Rows(0).Item("DestroyUploadTime") = tbDBInv007.Rows(0).Item("DestroyUploadTime")
                tbScrInv007.Rows(0).Item("DestroyReason") = tbDBInv007.Rows(0).Item("DestroyReason")
                tbScrInv007.Rows(0).Item("DepositMK") = tbDBInv007.Rows(0).Item("DepositMK")
                tbScrInv007.Rows(0).Item("DataType") = tbDBInv007.Rows(0).Item("DataType")
                tbScrInv007.Rows(0).Item("PrizeFile") = tbDBInv007.Rows(0).Item("PrizeFile")
            Else
                For Each rwscr007 As DataRow In tbScrInv007.Rows
                    Dim o As New Random

                    rwscr007.Item("CreateInvDate") = executeTime
                    rwscr007.Item("RandomNum") = Right("0000" & New Random().Next(1, 9999).ToString, 4)
                    rwscr007.Item("CompID") = LoginInfo.CompCode.ToString
                    tbScrInv007.Rows(0).Item("PrizeFile") = DBNull.Value
                Next
            End If
            For Each rwscr007 As DataRow In tbScrInv007.Rows
                rwscr007.Item("UptTime") = executeTime
                rwscr007.Item("UptEn") = LoginInfo.EntryName
            Next

        Catch ex As Exception
            Throw
        End Try
    End Sub
    'The function looks strange as  I copy grammar from delphi
    Private Function getCheckNo(ByVal InvID As String, ByVal InvDate As Date) As String
        Dim sL_Tmp As String = "", sL_Tmp1 = "", sL_Tmp2 = ""
        Dim nL_Length As Integer = 8, nL_CheckNo = 0
        Dim sL_Result As String = ""
        Try
            Using tb001 As DataTable = DAO.ExecQry(_DAL.QueryInv001, New Object() {LoginInfo.CompCode.ToString})
                Dim sI_SystemID As String = tb001.Rows(0).Item("CheckInvNum").ToString

                If String.IsNullOrEmpty(sI_SystemID) Then
                    Throw New Exception(Language.SystemIDNull)
                End If
                Dim sL_TargetInvID As String = Right(InvID, 8)
                Dim sL_TargetInvDate As String = sI_SystemID & (Integer.Parse(Date.Parse(InvDate).ToString("yyyy")) - 1911).ToString &
                                Date.Parse(InvDate).ToString("MMdd")
                Dim L_Tmp1(7) As Byte, L_Tmp2(7) As Byte, L_Tmp3(8) As Byte
                For i As Integer = 0 To nL_Length - 1
                    L_Tmp1(i) = Integer.Parse(sL_TargetInvID.Substring(i, 1))
                    L_Tmp2(i) = Integer.Parse(sL_TargetInvDate.Substring(i, 1))
                    L_Tmp3(i) = 0
                Next
                L_Tmp3(nL_Length) = 0
                Dim j As Integer = nL_Length
                For i As Integer = 0 To nL_Length - 1
                    sL_Tmp = (L_Tmp1(j - 1) * L_Tmp2(j - 1)).ToString
                    If sL_Tmp.Length = 2 Then
                        sL_Tmp1 = Left(sL_Tmp, 1)
                        sL_Tmp2 = Right(sL_Tmp, 1)
                    Else
                        sL_Tmp1 = "0"
                        sL_Tmp2 = Left(sL_Tmp, 1)
                    End If

                    L_Tmp3(i) = L_Tmp3(i) + Integer.Parse(sL_Tmp1)
                    L_Tmp3(i + 1) = L_Tmp3(i + 1) + Integer.Parse(sL_Tmp2)
                    j = j - 1
                Next

                For i As Integer = 0 To L_Tmp3.Length - 1
                    nL_CheckNo = nL_CheckNo + L_Tmp3(i)
                Next




                tb001.Dispose()
            End Using
            Return nL_CheckNo.ToString
        Catch ex As Exception
            Throw
        End Try

    End Function
    Public Function Save(ByVal editMode As EditMode, ByVal ds As DataSet) As RIAResult
        Dim result As New RIAResult With {.ResultBoolean = False, .ErrorCode = -1, .ErrorMessage = "init"}
        Dim cn As DbConnection = DAO.GetConn()
        Dim trans As DbTransaction = Nothing
        Dim aInvid As String = ds.Tables("INV007").Rows(0).Item("INVID")
        Dim tbDBInv007 As DataTable = DAO.ExecQry(_DAL.QueryINV007, New Object() {aInvid})
        Dim tbDBInv008 As DataTable = DAO.ExecQry(_DAL.QueryINV008, New Object() {aInvid})
        Dim tbDBInv008A As DataTable = DAO.ExecQry(_DAL.QueryINV008A, New Object() {aInvid})
        Dim tbCalculatedInv008 As DataTable = Nothing
        Dim tbInv001 As DataTable = DAO.ExecQry(_DAL.QueryInv001, New Object() {LoginInfo.CompCode.ToString})
        Dim aAutoCreateNum As Integer = 0
        Dim tbDel008A As DataTable = tbDBInv008A.Clone
        tbDel008A.Rows.Clear()

        If Not DBNull.Value.Equals(tbInv001.Rows(0).Item("AutoCreateNum")) Then
            aAutoCreateNum = Integer.Parse(tbInv001.Rows(0).Item("AutoCreateNum"))
        End If
        'order INV008
        ds.Tables("INV008").DefaultView.Sort = "NEWADD"
        ds.Tables("INV008").DefaultView.ToTable()
        DisplaceScrInv007(ds.Tables("INV007"), tbDBInv007)

        Dim blnTrans As Boolean = False
        Dim blnAutoClose As Boolean = False

        Try
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
                Select Case editMode
                    Case CableSoft.BLL.Utility.EditMode.Edit
                        aAction = "INV EDIT"
                    Case CableSoft.BLL.Utility.EditMode.Append
                        aAction = "INV ADD"
                    Case Else
                        aAction = "INV EDIT"
                End Select
                CableSoft.BLL.Utility.Utility.SetClientInfo(Me.DAO, LoginInfo.EntryId, aAction)
            End If

            'pick up delete data
            Dim rwDel008 As List(Of DataRow) = tbDBInv008.AsEnumerable.Where(Function(rwDB008)
                                                                                 Dim blnDele As Boolean = True
                                                                                 For i As Integer = 0 To ds.Tables("INV008").Rows.Count - 1
                                                                                     If rwDB008.Item("SEQ") = ds.Tables("INV008").Rows(i).Item("SEQ") AndAlso ds.Tables("INV008").Rows(i).Item("NEWADD") = 0 Then
                                                                                         blnDele = False
                                                                                         Exit For
                                                                                     End If
                                                                                 Next
                                                                                 Return blnDele
                                                                             End Function).ToList
            If rwDel008.Count > 0 Then

                For Each rw As DataRow In rwDel008

                    For i As Integer = 0 To tbDBInv008A.Rows.Count - 1
                        If tbDBInv008A.Rows(i).Item("SEQ") = rw.Item("SEQ") Then
                            If Not DBNull.Value.Equals(tbDBInv008A.Rows(i).Item("BILLID")) AndAlso Not DBNull.Value.Equals(tbDBInv008A.Rows(i).Item("BILLIDITEMNO")) Then
                                If tbDBInv008A.Rows(i).Item("BILLID") <> rw.Item("BILLID") OrElse tbDBInv008A.Rows(i).Item("BILLIDITEMNO") <> rw.Item("BILLIDITEMNO") Then
                                    If rw.Item("LinkToMIS").ToString.ToUpper = "Y" Then
                                        tbDel008A.Rows.Add(tbDBInv008A.Rows(i).ItemArray)
                                    End If
                                End If
                            End If
                            tbDBInv008A.Rows(i).Delete()
                        End If
                    Next
                    tbDBInv008A.AcceptChanges()
                Next
            End If

            'make sure INV008.CombCitemName and INV008.CombCitemCode have value
            For i As Integer = 0 To ds.Tables("INV008").Rows.Count - 1
                If Not DBNull.Value.Equals(ds.Tables("INV008").Rows(i).Item("CombCitemCode")) AndAlso Not DBNull.Value.Equals(ds.Tables("INV008").Rows(i).Item("CombCitemName")) Then

                Else
                    Using tbItemId As DataTable = DAO.ExecQry(_DAL.QueryItemId, New Object() {LoginInfo.CompCode.ToString, ds.Tables("INV008").Rows(i).Item("ItemID")})
                        If tbItemId.Rows.Count = 0 Then
                            ds.Tables("INV008").Rows(i).Item("CombCitemCode") = DBNull.Value
                            ds.Tables("INV008").Rows(i).Item("CombCitemName") = DBNull.Value
                        Else
                            If DBNull.Value.Equals(tbItemId.Rows(0).Item("ItemIdRef")) OrElse DBNull.Value.Equals(tbItemId.Rows(0).Item("ItemIdRefDesc")) Then
                                ds.Tables("INV008").Rows(i).Item("CombCitemCode") = DBNull.Value
                                ds.Tables("INV008").Rows(i).Item("CombCitemName") = DBNull.Value
                            Else
                                ds.Tables("INV008").Rows(i).Item("CombCitemCode") = tbItemId.Rows(0).Item("ItemIdRef")
                                ds.Tables("INV008").Rows(i).Item("CombCitemName") = tbItemId.Rows(0).Item("ItemIdRefDesc")
                            End If
                        End If

                    End Using
                End If
            Next

            'process inv008
            tbCalculatedInv008 = ProcessInv008(ds.Tables("INV008"), tbDBInv008A)
            'recalculate inv008
            ReCalculateInv008(tbCalculatedInv008, tbDBInv008A)
            'check available invnum
            If aAutoCreateNum > 0 Then
                If editMode = EditMode.Edit Then
                    If tbCalculatedInv008.Rows.Count > aAutoCreateNum Then
                        result.ErrorMessage = String.Format(Language.ExcessDetail, aAutoCreateNum)
                        If blnTrans Then
                            trans.Rollback()
                        End If
                        Return result
                    End If
                End If
                If editMode = EditMode.Append Then

                    result = checkCanCreate(ds.Tables("INV007"), tbCalculatedInv008, aAutoCreateNum)
                    If result.ResultBoolean = False Then Return result
                End If
            End If
            'calculate Inv007
            ReCalculateInv007(ds.Tables("INV007"), tbCalculatedInv008, tbDBInv008A, aAutoCreateNum)
            blnTrans = True
            'ready to update
            'update inv099 first

            Dim PREINVOICE As Integer = 0
            Select Case ds.Tables("INV007").Rows(0).Item("HowToCreate")
                Case "1"
                    PREINVOICE = 1
                Case "2"
                    PREINVOICE = 0
                Case "3"
                    PREINVOICE = 2
                Case Else
                    PREINVOICE = 2
            End Select
            Dim o As Integer = 0
            If editMode = EditMode.Append Then
                Using tbCurrentInv099 As DataTable = DAO.ExecQry(_DAL.QueryCurrentInv099, New Object() {
                                      LoginInfo.CompCode.ToString,
                                      Right(ds.Tables("INV007").Rows(0).Item("INVID").ToString, 8),
                                      getYearMonth(ds.Tables("INV007").Rows(0).Item("INVDATE")),
                                      ds.Tables("INV007").Rows(0).Item("INVID").ToString.Substring(0, 2)
                                 })
                    o = DAO.ExecNqry(_DAL.UpdateInv099, New Object() {
                                    ds.Tables("INV007").Rows.Count,
                                    ds.Tables("INV007").Rows(0).Item("INVDATE"),
                                    LoginInfo.CompCode.ToString,
                                    tbCurrentInv099.Rows(0).Item("YearMonth"),
                                    tbCurrentInv099.Rows(0).Item("Prefix"),
                                    tbCurrentInv099.Rows(0).Item("StartNum")
                              })
                    o = DAO.ExecNqry(_DAL.UpdateIn099Useful, New Object() {
                                            LoginInfo.CompCode.ToString,
                                            tbCurrentInv099.Rows(0).Item("YearMonth"),
                                            tbCurrentInv099.Rows(0).Item("Prefix"),
                                            tbCurrentInv099.Rows(0).Item("StartNum")
                                     })
                End Using
            End If


            'delete inv007 if data exists
            If tbDBInv007.Rows.Count > 0 Then
                DAO.ExecNqry(_DAL.DeleteInv007DataByInvId, New Object() {tbDBInv007.Rows(0).Item("INVID"), LoginInfo.CompCode.ToString})
                DAO.ExecNqry(_DAL.DeleteINV008ByInvId, New Object() {tbDBInv007.Rows(0).Item("INVID")})
                DAO.ExecNqry(_DAL.DeleteINV008AByInvId, New Object() {tbDBInv007.Rows(0).Item("INVID")})
            End If
            'insert all screen inv007
            For i As Integer = 0 To ds.Tables("INV007").Rows.Count - 1
                Dim aInsertAndValues As Dictionary(Of String, Object) =
                        GetInsertInvDataSQL(ds.Tables("INV007"), tbDBInv007, i, "INV007")
                DAO.ExecNqry(aInsertAndValues("SQL"), CType(aInsertAndValues("VALUES"), ArrayList).ToArray)
                aInsertAndValues.Clear()
            Next
            'insert all screen inv008
            If tbDBInv008.Columns.Contains("newadd") Then tbDBInv008.Columns.Remove("newadd")
            If tbDBInv008.Columns.Contains("MergeFlag") Then tbDBInv008.Columns.Remove("MergeFlag")
            For i As Integer = 0 To tbCalculatedInv008.Rows.Count - 1
                Dim aInsertAndValues As Dictionary(Of String, Object) =
                        GetInsertInvDataSQL(tbCalculatedInv008, tbDBInv008, i, "INV008")
                DAO.ExecNqry(aInsertAndValues("SQL"), CType(aInsertAndValues("VALUES"), ArrayList).ToArray)
                aInsertAndValues.Clear()
                'update bill
                If tbCalculatedInv008.Rows(i).Item("LinkToMIS").ToString.ToUpper() = "Y" Then
                    If Not DBNull.Value.Equals(tbCalculatedInv008.Rows(i).Item("BillID")) AndAlso Not DBNull.Value.Equals(tbCalculatedInv008.Rows(i).Item("BillIDItemNo")) Then
                        DAO.ExecNqry(_DAL.UpdateSOBill("SO033"), New Object() {
                                         tbCalculatedInv008.Rows(i).Item("INVID"),
                                        PREINVOICE, ds.Tables("INV007").Rows(0).Item("INVDATE"),
                                        executeTime, ds.Tables("INV007").Rows(0).Item("InvUseId"),
                                        ds.Tables("INV007").Rows(0).Item("InvUseDesc"),
                                        tbCalculatedInv008.Rows(i).Item("BillID"),
                                        tbCalculatedInv008.Rows(i).Item("BillIDItemNo")
                        })
                        DAO.ExecNqry(_DAL.UpdateSOBill("SO034"), New Object() {
                                        tbCalculatedInv008.Rows(i).Item("INVID"),
                                       PREINVOICE, ds.Tables("INV007").Rows(0).Item("INVDATE"),
                                       executeTime, ds.Tables("INV007").Rows(0).Item("InvUseId"),
                                       ds.Tables("INV007").Rows(0).Item("InvUseDesc"),
                                        tbCalculatedInv008.Rows(i).Item("BillID"),
                                        tbCalculatedInv008.Rows(i).Item("BillIDItemNo")
                       })


                    End If
                    'upd inv008a data to so
                    For j As Integer = 0 To tbDBInv008A.Rows.Count - 1
                        If (tbDBInv008A.Rows(j).Item("SEQ") = tbCalculatedInv008.Rows(i).Item("SEQ")) Then
                            'Dim aInsertAndValues8A As Dictionary(Of String, Object) =
                            '   GetInsertInvDataSQL(tbDBInv008A, tbDBInv008A, j, "INV008A")
                            'DAO.ExecNqry(aInsertAndValues8A("SQL"), CType(aInsertAndValues8A("VALUES"), ArrayList).ToArray)
                            'aInsertAndValues8A.Clear()

                            If Not DBNull.Value.Equals(tbDBInv008A.Rows(j).Item("BillID")) AndAlso Not DBNull.Value.Equals(tbDBInv008A.Rows(j).Item("BillIDItemNo")) Then
                                DAO.ExecNqry(_DAL.UpdateSOBill("SO033"), New Object() {
                                                 tbDBInv008A.Rows(j).Item("INVID"),
                                                PREINVOICE, ds.Tables("INV007").Rows(0).Item("INVDATE"),
                                                executeTime, ds.Tables("INV007").Rows(0).Item("InvUseId"),
                                                ds.Tables("INV007").Rows(0).Item("InvUseDesc"),
                                                tbDBInv008A.Rows(j).Item("BillID"),
                                                tbDBInv008A.Rows(j).Item("BillIDItemNo")
                                })
                                DAO.ExecNqry(_DAL.UpdateSOBill("SO034"), New Object() {
                                                tbDBInv008A.Rows(j).Item("INVID"),
                                               PREINVOICE, ds.Tables("INV007").Rows(0).Item("INVDATE"),
                                               executeTime, ds.Tables("INV007").Rows(0).Item("InvUseId"),
                                               ds.Tables("INV007").Rows(0).Item("InvUseDesc"),
                                                tbDBInv008A.Rows(j).Item("BillID"),
                                                tbDBInv008A.Rows(j).Item("BillIDItemNo")
                               })


                            End If
                        End If




                    Next


                End If


            Next
            'insert 8a to db
            For i As Integer = 0 To tbDBInv008A.Rows.Count - 1
                Dim aInsertAndValues8A As Dictionary(Of String, Object) =
                               GetInsertInvDataSQL(tbDBInv008A, tbDBInv008A, i, "INV008A")
                DAO.ExecNqry(aInsertAndValues8A("SQL"), CType(aInsertAndValues8A("VALUES"), ArrayList).ToArray)
                aInsertAndValues8A.Clear()
            Next
            'restore SO.bill
            If rwDel008.Count > 0 Then
                For Each rw As DataRow In rwDel008
                    If Not DBNull.Value.Equals(rw.Item("BILLID")) AndAlso Not DBNull.Value.Equals(rw.Item("BILLIDITEMNO")) AndAlso rw.Item("LinkToMIS").ToString.ToUpper = "Y" Then
                        DAO.ExecNqry(_DAL.ReStoreSOBillByINV("SO033", "INV008"), New Object() {rw("BILLID"), rw("BILLIDITEMNO")})
                        DAO.ExecNqry(_DAL.ReStoreSOBillByINV("SO034", "INV008"), New Object() {rw("BILLID"), rw("BILLIDITEMNO")})
                    End If



                    'Using tb008A As DataTable = DAO.ExecQry(_DAL.QueryINV008ASingle,
                    '                                       New Object() {rw.Item("INVID"), rw.Item("SEQ")})
                    '    DAO.ExecNqry(_DAL.ReStoreSOBillByINV("SO033", "INV008"), New Object() {tb008A.Rows(0).Item("BILLID"), tb008A.Rows(0).Item("BILLIDITEM")})
                    '    DAO.ExecNqry(_DAL.ReStoreSOBillByINV("SO034", "INV008"), New Object() {tb008A.Rows(0).Item("BILLID"), tb008A.Rows(0).Item("BILLIDITEM")})

                    'End Using

                Next
            End If
            If tbDel008A.Rows.Count > 0 Then
                For Each rw As DataRow In tbDel008A.Rows
                    DAO.ExecNqry(_DAL.ReStoreSOBillByINV("SO033", "INV008"), New Object() {rw("BILLID"), rw("BILLIDITEMNO")})
                    DAO.ExecNqry(_DAL.ReStoreSOBillByINV("SO034", "INV008"), New Object() {rw("BILLID"), rw("BILLIDITEMNO")})
                Next
            End If

            result.ErrorCode = 0
            result.ResultBoolean = True
            result.ErrorMessage = Nothing
            Dim okInvId As String = Nothing
            For Each rw007 As DataRow In ds.Tables("INV007").Rows
                If String.IsNullOrEmpty(okInvId) Then
                    okInvId = rw007.Item("INVID")
                Else
                    okInvId = okInvId & "," & rw007.Item("INVID")
                End If
            Next
            If editMode = EditMode.Append Then
                result.ResultXML = okInvId
                result.Message = String.Format(Language.TotalCreate, okInvId, ds.Tables("INV007").Rows.Count)
            Else
                result.ResultXML = Nothing
            End If


            If blnTrans Then
                trans.Commit()
            End If

        Catch ex As Exception
            If blnTrans Then
                trans.Rollback()
            End If

            Throw ex
        Finally
            If tbCalculatedInv008 IsNot Nothing Then
                tbCalculatedInv008.Dispose()
                tbCalculatedInv008 = Nothing
            End If
            If tbDBInv007 IsNot Nothing Then
                tbDBInv007.Dispose()
                tbDBInv007 = Nothing
            End If
            If tbDBInv008 IsNot Nothing Then
                tbDBInv008.Dispose()
                tbDBInv008 = Nothing
            End If
            If tbDBInv008A IsNot Nothing Then
                tbDBInv008A.Dispose()
                tbDBInv008A = Nothing
            End If
            If tbInv001 IsNot Nothing Then
                tbInv001.Dispose()
                tbInv001 = Nothing
            End If

        End Try
        Return result
    End Function

    Private Function GetInsertInvDataSQL(ByVal tbScr As DataTable, tbDB As DataTable, ByVal rwIndex As Integer, ByVal tbName As String) As Dictionary(Of String, Object)
        'Dim result As String = Nothing
        Dim aField As String = Nothing
        Dim aValue As String = Nothing
        Dim aAryValue As New ArrayList()
        Dim debugIndex As Integer = 0
        Dim aInsertAndValue As New Dictionary(Of String, Object)
        aInsertAndValue.Add("SQL", Nothing)
        aInsertAndValue.Add("VALUES", Nothing)
        Dim iSign As Integer = 0
        For Each colDB As DataColumn In tbDB.Columns
            If tbScr.Columns.Contains(colDB.ColumnName) Then

                If String.IsNullOrEmpty(aField) Then
                    aField = colDB.ColumnName
                Else
                    aField = aField & "," & colDB.ColumnName
                End If
                If String.IsNullOrEmpty(aValue) Then
                    aValue = _DAL.Sign & iSign
                Else
                    aValue = aValue & "," & _DAL.Sign & iSign
                End If
                aAryValue.Add(tbScr.Rows(rwIndex).Item(colDB.ColumnName))
                iSign += 1
                debugIndex += 1
            End If
        Next
        aInsertAndValue("SQL") = "Insert Into " & tbName & " (" & aField & ") Values (" & aValue & ") "
        aInsertAndValue("VALUES") = aAryValue
        Return aInsertAndValue
    End Function
    Private Sub ReCalculateInv007(ByVal tbInv007 As DataTable, ByVal tbInv008 As DataTable,
                                  ByVal tbInv008A As DataTable,
                                  ByVal AutoCreateNum As Integer)
        Try
            Dim aSEQ As Integer = 1
            Dim aOldInv008SEQ As Integer = 0
            Dim blnHasChange As Boolean = False
            Dim aFaciSNo008 As String = Nothing
            Dim MainId As String = tbInv007.Rows(0).Item("INVID")
            Dim changeInvId = Nothing
            If Not DBNull.Value.Equals(tbInv007.Rows(0).Item("MainINVID")) Then
                MainId = tbInv007.Rows(0).Item("MainINVID")
            End If
            tbInv007.Rows(0).Item("MainINVID") = MainId
            If AutoCreateNum = 0 Then AutoCreateNum = 9999
            For i As Integer = 0 To tbInv008.Rows.Count - 1

                aOldInv008SEQ = tbInv008.Rows(i).Item("SEQ")
                'If (((i + 1) / AutoCreateNum) Mod (AutoCreateNum + 1)) = 0 Then
                If (i + 1) Mod (AutoCreateNum + 1) = 0 Then
                    aSEQ = 1

                    blnHasChange = True
                    'append new Inv007
                    Dim rwNew As DataRow = tbInv007.NewRow
                    rwNew.ItemArray = tbInv007.Copy.Rows(0).ItemArray
                    changeInvId = Left(rwNew.Item("INVID"), 2) & Right("00000000" & Integer.Parse(Right(rwNew.Item("INVID"), 8) + 1), 8)
                    rwNew.Item("INVID") = changeInvId
                    rwNew.Item("MainInvID") = MainId

                    tbInv007.Rows.Add(rwNew.ItemArray)
                    tbInv007.AcceptChanges()
                End If
                'update inv008 and inv008a relation
                '---------------------------------------------------------
                If blnHasChange Then
                    tbInv008.Rows(i).Item("SEQ") = aSEQ
                    tbInv008.Rows(i).Item("INVID") = changeInvId
                    For Each rw008A As DataRow In tbInv008A.Rows
                        If rw008A.Item("SEQ") = aOldInv008SEQ Then
                            rw008A.Item("SEQ") = aSEQ
                            rw008A.Item("INVID") = changeInvId
                        End If
                    Next
                End If
                '---------------------------------------------------------
                'If (Not DBNull.Value.Equals(tbInv008.Rows(i).Item("FaciSNo"))) Then
                '    If String.IsNullOrEmpty(aFaciSNo008) Then
                '        aFaciSNo008 = tbInv008.Rows(i).Item("FaciSNo")
                '    Else
                '        aFaciSNo008 = aFaciSNo008 & "," & tbInv008.Rows(i).Item("FaciSNo")
                '    End If
                'End If
                aSEQ += 1

            Next
            For Each rwInv007 As DataRow In tbInv007.Rows
                aFaciSNo008 = Nothing
                rwInv007.Item("TaxAmount") = tbInv008.AsEnumerable.Sum(Function(rw008)
                                                                           If rw008.Item("INVID") = rwInv007.Item("INVID") Then
                                                                               Return rw008.Item("TaxAmount")
                                                                           Else
                                                                               Return 0
                                                                           End If
                                                                       End Function)
                rwInv007.Item("SaleAmount") = tbInv008.AsEnumerable.Sum(Function(rw008)
                                                                            If rw008.Item("INVID") = rwInv007.Item("INVID") Then
                                                                                Return rw008.Item("SaleAmount")
                                                                            Else
                                                                                Return 0
                                                                            End If
                                                                        End Function)
                rwInv007.Item("InvAmount") = tbInv008.AsEnumerable.Sum(Function(rw008)
                                                                           If rw008.Item("INVID") = rwInv007.Item("INVID") Then
                                                                               Return rw008.Item("TotalAmount")
                                                                           Else
                                                                               Return 0
                                                                           End If
                                                                       End Function)
                rwInv007.Item("CheckNo") = getCheckNo(rwInv007.Item("INVID"), rwInv007.Item("INVDATE"))
                For Each rwFaciSNo As DataRow In tbInv008.AsEnumerable.Where(Function(rw008)
                                                                                 If rw008.Item("INVID") = rwInv007.Item("INVID") Then
                                                                                     Return True
                                                                                 Else
                                                                                     Return False
                                                                                 End If
                                                                             End Function).ToArray

                    If Not DBNull.Value.Equals(rwFaciSNo.Item("FacisNo")) Then
                        If String.IsNullOrEmpty(aFaciSNo008) Then
                            aFaciSNo008 = rwFaciSNo.Item("FacisNo")
                        Else
                            aFaciSNo008 = aFaciSNo008 & "," & rwFaciSNo.Item("FacisNo")
                        End If
                    End If
                Next
                If Not String.IsNullOrEmpty(aFaciSNo008) Then
                    If DBNull.Value.Equals(rwInv007.Item("Memo1")) Then
                        rwInv007.Item("Memo1") = "設備:" & aFaciSNo008
                    Else
                        rwInv007.Item("Memo1") = rwInv007.Item("Memo1") & "設備:" & aFaciSNo008
                    End If
                End If


                rwInv007.Item("MainTaxAmount") = tbInv008.AsEnumerable.Sum(Function(rw008)
                                                                               Return rw008.Item("TaxAmount")

                                                                           End Function)
                rwInv007.Item("MainSaleAmount") = tbInv008.AsEnumerable.Sum(Function(rw008)
                                                                                Return rw008.Item("SaleAmount")

                                                                            End Function)
                rwInv007.Item("MainInvAmount") = tbInv008.AsEnumerable.Sum(Function(rw008)
                                                                               Return rw008.Item("TotalAmount")
                                                                           End Function)
            Next

        Catch ex As Exception
            Throw ex
        End Try

    End Sub
    Private Sub ReCalculateInv008(tbInv008 As DataTable, tbInv008A As DataTable)
        Try
            Dim aOldSEQ As Integer = 0
            Dim Quantity As Integer = 0, UnitPrice As Integer = 0, TaxAmount As Integer = 0, TotalAmount As Integer = 0
            Dim SaleAmount As Integer = 0
            Dim StartDate As Object = Nothing, EndDate As Object = Nothing
            Dim FacisNo As String = Nothing, AccountNo As String = Nothing, SmartCardNo As String = Nothing, CMMac As String = Nothing
            For i As Integer = 0 To tbInv008.Rows.Count - 1
                aOldSEQ = tbInv008.Rows(i).Item("SEQ")

                tbInv008.Rows(i).Item("SEQ") = i + 1
                For Each rw008a As DataRow In tbInv008A.Rows
                    If rw008a.Item("SEQ") = aOldSEQ Then
                        tbInv008.Rows(i).Item("CombCitemCode") = DBNull.Value
                        tbInv008.Rows(i).Item("CombCitemName") = DBNull.Value
                        Quantity = 1
                        If Not DBNull.Value.Equals(rw008a.Item("FacisNo")) Then
                            FacisNo = rw008a.Item("FacisNo")
                        End If
                        If Not DBNull.Value.Equals(rw008a.Item("AccountNo")) Then
                            AccountNo = rw008a.Item("AccountNo")
                        End If
                        If Not DBNull.Value.Equals(rw008a.Item("SmartCardNo")) Then
                            SmartCardNo = rw008a.Item("SmartCardNo")
                        End If
                        If Not DBNull.Value.Equals(rw008a.Item("CMMac")) Then
                            CMMac = rw008a.Item("CMMac")
                        End If
                        UnitPrice += rw008a.Item("UnitPrice")
                        TaxAmount += rw008a.Item("TaxAmount")
                        TotalAmount += rw008a.Item("TotalAmount")
                        SaleAmount += rw008a.Item("SaleAmount")
                        If Not DBNull.Value.Equals(rw008a.Item("StartDate")) Then
                            If StartDate Is Nothing Then
                                StartDate = Date.Parse(rw008a.Item("StartDate"))
                            Else
                                If Date.Parse(StartDate) > Date.Parse(rw008a.Item("StartDate")) Then
                                    StartDate = Date.Parse(rw008a.Item("StartDate"))
                                End If
                            End If
                        End If
                        If Not DBNull.Value.Equals(rw008a.Item("EndDate")) Then
                            If EndDate Is Nothing Then
                                EndDate = Date.Parse(rw008a.Item("EndDate"))
                            Else
                                If Date.Parse(EndDate) < Date.Parse(rw008a.Item("EndDate")) Then
                                    EndDate = Date.Parse(rw008a.Item("EndDate"))
                                End If
                            End If
                        End If
                        rw008a.Item("SEQ") = i + 1
                    End If
                Next
                If Quantity <> 0 Then
                    tbInv008.Rows(i).Item("Quantity") = Quantity
                    tbInv008.Rows(i).Item("UnitPrice") = UnitPrice
                    tbInv008.Rows(i).Item("TaxAmount") = TaxAmount
                    tbInv008.Rows(i).Item("TotalAmount") = TotalAmount
                    tbInv008.Rows(i).Item("SaleAmount") = SaleAmount
                End If
                If Not String.IsNullOrEmpty(FacisNo) Then
                    tbInv008.Rows(i).Item("FacisNo") = FacisNo
                End If
                If Not String.IsNullOrEmpty(AccountNo) Then
                    tbInv008.Rows(i).Item("AccountNo") = AccountNo
                End If
                If Not String.IsNullOrEmpty(SmartCardNo) Then
                    tbInv008.Rows(i).Item("SmartCardNo") = SmartCardNo
                End If
                If Not String.IsNullOrEmpty(CMMac) Then
                    tbInv008.Rows(i).Item("CMMac") = CMMac
                End If
                tbInv008.Rows(i).Item("UptTime") = executeTime
                tbInv008.Rows(i).Item("UptEn") = LoginInfo.EntryName
                Quantity = 0
                UnitPrice = 0
                TaxAmount = 0
                TotalAmount = 0
                SaleAmount = 0
                FacisNo = Nothing
                AccountNo = Nothing
                SmartCardNo = Nothing
                CMMac = Nothing
            Next
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Function ProcessInv008(tbScrInv008 As DataTable, tbDBInv008A As DataTable) As DataTable
        Try
            Dim tbInv008Clone As DataTable = tbScrInv008.Clone
            tbInv008Clone.Clear()


            For i As Integer = 0 To tbScrInv008.Rows.Count - 1
                'find scrinv008 addnew and ignore old data (add directly)

                If Not DBNull.Value.Equals(tbScrInv008.Rows(i).Item("CombCitemCode")) AndAlso
                        tbScrInv008.Rows(i).Item("NEWADD") = 1 Then
                    'find new create inv008 whether has merge data
                    '-----------------------------------------------------------------------------------------------------------
                    'add to inv008a
                    Dim newInv008Arw As DataRow = tbDBInv008A.NewRow
                    newInv008Arw.Item("ItemIdRef") = tbScrInv008.Rows(i).Item("CombCitemCode")
                    newInv008Arw.Item("InvID") = tbScrInv008.Rows(i).Item("INVID")
                    newInv008Arw.Item("BillID") = tbScrInv008.Rows(i).Item("BillID")
                    newInv008Arw.Item("BillIDItemNo") = tbScrInv008.Rows(i).Item("BillIDItemNo")
                    newInv008Arw.Item("SEQ") = tbScrInv008.Rows(i).Item("SEQ")
                    newInv008Arw.Item("StartDate") = tbScrInv008.Rows(i).Item("StartDate")
                    newInv008Arw.Item("EndDate") = tbScrInv008.Rows(i).Item("EndDate")
                    newInv008Arw.Item("ItemID") = tbScrInv008.Rows(i).Item("ItemID")
                    newInv008Arw.Item("Description") = tbScrInv008.Rows(i).Item("Description")
                    newInv008Arw.Item("Quantity") = tbScrInv008.Rows(i).Item("Quantity")
                    newInv008Arw.Item("UnitPrice") = tbScrInv008.Rows(i).Item("UnitPrice")
                    newInv008Arw.Item("TaxAmount") = tbScrInv008.Rows(i).Item("TaxAmount")
                    newInv008Arw.Item("SaleAmount") = tbScrInv008.Rows(i).Item("SaleAmount")
                    newInv008Arw.Item("TotalAmount") = tbScrInv008.Rows(i).Item("TotalAmount")
                    newInv008Arw.Item("ServiceType") = tbScrInv008.Rows(i).Item("ServiceType")
                    newInv008Arw.Item("FacisNo") = tbScrInv008.Rows(i).Item("FacisNo")
                    newInv008Arw.Item("AccountNo") = tbScrInv008.Rows(i).Item("AccountNo")
                    newInv008Arw.Item("SmartCardNo") = tbScrInv008.Rows(i).Item("SmartCardNo")
                    newInv008Arw.Item("CMMac") = tbScrInv008.Rows(i).Item("CMMac")

                    '-----------------------------------------------------------------------------------------------------------
                    '
                    If tbInv008Clone.AsEnumerable.Count(Function(ByVal rwClone As DataRow)
                                                            Return rwClone.Item("ItemID") = tbScrInv008.Rows(i).Item("CombCitemCode")
                                                        End Function) = 0 Then
                        Dim rw008 As DataRow = tbScrInv008.NewRow
                        rw008.ItemArray = tbScrInv008.Rows(i).ItemArray
                        rw008.Item("ItemId") = tbScrInv008.Rows(i).Item("CombCitemCode")
                        rw008.Item("Description") = tbScrInv008.Rows(i).Item("CombCitemName")
                        tbInv008Clone.Rows.Add(rw008.ItemArray)

                    Else
                        newInv008Arw.Item("SEQ") = tbInv008Clone.AsEnumerable.First(Function(rw008)
                                                                                        Return rw008.Item("ItemID") = tbScrInv008.Rows(i).Item("CombCitemCode")
                                                                                    End Function).Item("SEQ")
                        'do nothing as inv008 find the same itemid
                        ' tbInv008Clone.Rows.Add(tbScrInv008.Rows(i).ItemArray)
                    End If
                    tbDBInv008A.Rows.Add(newInv008Arw)
                    tbDBInv008A.AcceptChanges()

                Else
                    'original inv008 data 
                    tbInv008Clone.Rows.Add(tbScrInv008.Rows(i).ItemArray)

                End If
            Next
            Return tbInv008Clone.Copy
        Catch ex As Exception
            Throw
        End Try
    End Function
    Public Function QuerySOCustByQuery(ByVal QueryIndex As Integer, ByVal QueryText As String, ByVal RefNo As Integer) As DataSet
        Dim ds As New DataSet
        Dim strCustId As String = "-1"
        Try
            If QueryIndex = 1 AndAlso RefNo = 3 Then QueryText = Integer.Parse(QueryText)
            Dim tbSO138 As DataTable = Nothing
            If RefNo = 3 Then
                tbSO138 = DAO.ExecQry(_DAL.QuerySOCustByQuery(QueryIndex), New Object() {QueryText, misCompCode}).Copy
            Else
                tbSO138 = DAO.ExecQry(_DAL.QueryINVCustInfo(QueryIndex), New Object() {LoginInfo.CompCode.ToString, QueryText}).Copy
                tbSO138.TableName = "SO138"
            End If

            For Each rw As DataRow In tbSO138.Rows
                strCustId = String.Format("{0},{1}", strCustId, rw.Item("CustID"))
            Next
            If RefNo = 4 Then
                ds.Tables.Add(tbSO138)
                Dim tbQuery As DataTable = DAO.ExecQry(_DAL.QueryCustWhere).Copy
                tbQuery.TableName = "QUERY"
                ds.Tables.Add(tbQuery)
            Else
                ds = QuerySOCustInfo(strCustId, "-X")
            End If



        Catch ex As Exception
            Throw ex
        End Try


        Return ds.Copy
    End Function
    Public Function QuerySOCustInfo(ByVal custid As String, ByVal existsBill As String) As DataSet
        Dim ds As New DataSet()
        Try
            Dim invseqno As String = "-1"
            Dim invServiceTypestr As String = "'X'"
            Dim tbSO138 As DataTable = Nothing
            Dim A_CarrierType As String = Nothing
            For Each o As String In existsBill.Split(", ")
                o = Replace(o, "'", "")
                If o.Length > 3 Then
                    Using tb As DataTable = DAO.ExecQry(_DAL.QueryBillInvseqNo(),
                                New Object() {Left(o, o.Length - 1), Right(o, 1)})
                        If tb.Rows.Count > 0 Then
                            If Not DBNull.Value.Equals(tb.Rows(0).Item(0)) Then
                                invseqno = tb.Rows(0).Item(0)
                            End If

                            Exit For
                        End If
                    End Using

                End If
            Next
            Using tb As DataTable = DAO.ExecQry(_DAL.QueryInv001, New Object() {LoginInfo.CompCode.ToString()})
                For Each rw As DataRow In tb.Rows
                    If Not IsDBNull(rw.Item("ServiceTypeStr")) Then
                        For Each o As String In rw.Item("ServiceTypeStr").ToString.Split(",")
                            invServiceTypestr = String.Format("{0},'{1}'", invServiceTypestr, o)
                        Next
                    End If
                    If Not DBNull.Value.Equals(rw.Item("A_CarrierType")) Then
                        A_CarrierType = rw.Item("A_CarrierType")
                    End If
                Next
                tb.Dispose()
            End Using
            Select Case pNewFlow
                Case 0
                    tbSO138 = DAO.ExecQry(_DAL.QueryOldSOCustInfo(custid, invServiceTypestr, A_CarrierType)).Copy
                Case Else
                    tbSO138 = DAO.ExecQry(_DAL.QuerySOCustInfo(invseqno, custid), New Object() {
                        misCompCode}).Copy
            End Select
            'Dim tbSO138 As DataTable = DAO.ExecQry(_DAL.QuerySOCustInfo(invseqno, custid), New Object() {
            '     LoginInfo.CompCode}).Copy

            Dim tbQuery As DataTable = DAO.ExecQry(_DAL.QueryCustWhere).Copy
            tbQuery.TableName = "QUERY"
            tbSO138.TableName = "SO138"
            ds.Tables.Add(tbSO138)
            ds.Tables.Add(tbQuery)
            'ds.Tables.Add(QuerySOCustId(custid).Tables("SO001").Copy)
        Catch ex As Exception
            Throw ex

        End Try


        Return ds.Copy

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
