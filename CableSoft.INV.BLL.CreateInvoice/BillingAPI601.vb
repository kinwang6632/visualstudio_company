Imports System.Data.Common
Imports CableSoft.BLL.BillingAPI
Imports CableSoft.BLL.Utility
Public Class BillingAPI601
    Inherits BLLBasic
    Implements IDisposable, CableSoft.BLL.BillingAPI.IBillingAPI
    Private _DAL As New BillingAPI601DALMultiDB(Me.LoginInfo.Provider)
    Private Language As New CableSoft.BLL.Language.SO61.BillingAPI601Language
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
        Dim aBillCollection As List(Of String) = Nothing
        Dim aInvID As Object = InData.Tables("Inv").Rows(0).Item("InvID")
        Dim aCompCode As Integer = InData.Tables("Main").Rows(0).Item("Compcode")
        Dim aObsoleteId As Object = InData.Tables("Inv").Rows(0).Item("ObsoleteId")
        Dim aHowToCreate As Integer = 0
        Dim tbMainInv007 As DataTable = Nothing
        Dim result As New RIAResult With {.ResultBoolean = False, .ErrorCode = -1, .ErrorMessage = "NO"}
        Dim FNowDate = Date.Now


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
        'Using EmpName As DataTable = DAO.ExecQry(_DAL.GetEmpName(aCusOwner), New Object() {InData.Tables("Inv").Rows(0).Item("Upden")})
        '    Me.LoginInfo.EntryId = InData.Tables("Inv").Rows(0).Item("Upden")
        '    Me.LoginInfo.EntryName = EmpName.Rows(0).Item("EmpName")
        'End Using
        Try
            If Not DBNull.Value.Equals(aCusComp) AndAlso Not DBNull.Value.Equals(aCusOwner) Then
                aLinkToMIS = True
            End If
            result = CanDropInv(aInvID, aCompCode)
            If result.ResultBoolean = False Then
                Return result
            End If
            tbMainInv007 = DAO.ExecQry(_DAL.QueryInv007, New Object() {aCompCode, aInvID})
            aHowToCreate = Integer.Parse(tbMainInv007.Rows(0).Item("HowToCreate").ToString())
            For Each row As DataRow In tbMainInv007.Rows
                DAO.ExecNqry(_DAL.DropInv007(), New Object() {aObsoleteId, aObsoleteId, _
                                                              Me.LoginInfo.EntryName, _
                                                              aCompCode, row.Item("InvId")})
                For i As Integer = 0 To 1
                    DAO.ExecNqry(_DAL.InsInv024(i), New Object() {aCompCode, row.Item("InvId")})
                    If aLinkToMIS Then
                        If aHowToCreate = 1 OrElse aHowToCreate = 2 OrElse aHowToCreate = 3 Then
                            DAO.ExecNqry(_DAL.updSO033SO034(aCusOwner, i), New Object() {CableSoft.BLL.Utility.DateTimeUtility.GetDTString(FNowDate),
                                FNowDate, Me.LoginInfo.EntryName, row.Item("INVID"), aCusComp})
                        End If

                    End If

                Next
                
            Next

            
            'If aLinkToMIS Then


            '    If aHowToCreate = 1 OrElse aHowToCreate = 2 OrElse aHowToCreate = 3 Then
            '        If aBillCollection Is Nothing Then aBillCollection = New List(Of String)
            '        For Each row As DataRow In tbMainInv007.Rows
            '            Dim o As List(Of String) = CatchBillData(row.Item("InvId"), aCompCode)
            '            If o IsNot Nothing AndAlso o.Count > 0 Then
            '                aBillCollection.AddRange(o.ToArray)
            '            End If
            '            'aBillCollection = CatchBillData(aInvID, aCompCode)
            '            If aBillCollection IsNot Nothing AndAlso aBillCollection.Count > 0 Then


            '            End If
            '        Next


            '    End If



            'End If

            If blnAutoClose Then
                If result.ResultBoolean Then
                    trans.Commit()
                Else
                    trans.Rollback()
                End If
            End If
            result.ResultBoolean = True
            result.ErrorCode = 1
            result.ErrorMessage = String.Format(Language.resultMsg, aInvID)
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
            If aBillCollection IsNot Nothing Then
                aBillCollection.Clear()
                aBillCollection = Nothing
            End If
            If tbMainInv007 IsNot Nothing Then
                tbMainInv007.Dispose()
                tbMainInv007 = Nothing
            End If

        End Try

    End Function
   
    Private Function CanDropInv(ByVal aInvId As String, ByVal aCompCode As String) As RIAResult
        Dim result As New RIAResult
        result.ResultBoolean = True
        Using tbInv007 As DataTable = DAO.ExecQry(_DAL.QueryInv007(), New Object() {aCompCode, aInvId})
            If tbInv007.Rows.Count = 0 Then
                result.ResultBoolean = False
                result.ErrorCode = 631
                result.ErrorMessage = String.Format(Language.noInvid, aInvId)
                Return result
            End If
            If tbInv007.Rows(0).Item("ISOBSOLETE").ToString = "Y" Then
                result.ResultBoolean = False
                result.ErrorCode = -1
                result.ErrorMessage = String.Format(Language.hasCancel, aInvId)
                Return result
            End If
            If Integer.Parse(DAO.ExecSclr(_DAL.QueryInv018(), New Object() {aCompCode, _
                                                                            Date.Parse(tbInv007.Rows(0).Item("INVDATE")).ToString("yyyyMM")})) > 0 Then
                result.ResultBoolean = False
                result.ErrorCode = -1
                result.ErrorMessage = String.Format(Language.hasLock,
                                                    Date.Parse(tbInv007.Rows(0).Item("INVDATE")).ToString("yyyy/MM"))
                Return result
            End If
            If Integer.Parse(DAO.ExecSclr(_DAL.QueryCountINV014(), New Object() {aCompCode, aInvId})) > 0 Then
                result.ResultBoolean = False
                result.ErrorCode = -1
                result.ErrorMessage = String.Format(Language.hasAllowance, aInvId)
                Return result
            End If
            tbInv007.Dispose()
        End Using
        Return result
    End Function
   
    Private Function CatchBillData(ByVal aInvId As String, ByVal aCompCode As String) As List(Of String)
        Dim ret As New List(Of String)
        Using tbINV008 As DataTable = DAO.ExecQry(_DAL.QueryInv008(), New Object() {"1", 0, aCompCode, aInvId})
            For Each row As DataRow In tbINV008.Rows
                If DBNull.Value.Equals(row.Item("BILLIDITEMNO")) OrElse Integer.Parse(row.Item("BILLIDITEMNO")) < 0 Then
                    ret.Add(row.Item("BILLID") & "X")
                Else
                    ret.Add(row.Item("BILLID") & row.Item("BILLIDITEMNO"))
                End If
                Using tbInv008A As DataTable = DAO.ExecQry(_DAL.QueryInv008A, New Object() {aInvId, row("SEQ")})
                    For Each row8A As DataRow In tbInv008A.Rows
                        If DBNull.Value.Equals(row8A.Item("BILLIDITEMNO")) Then
                            ret.Add(row8A.Item("BILLID") & "X")
                        Else
                            ret.Add(row8A.Item("BILLID") & row8A.Item("BILLIDITEMNO"))
                        End If
                    Next
                    tbInv008A.Dispose()
                End Using


            Next
            tbINV008.Dispose()
        End Using
        
       

        Return ret.Distinct().ToList
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
