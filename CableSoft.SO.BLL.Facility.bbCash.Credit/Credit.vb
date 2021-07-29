Imports CableSoft.BLL.Utility
Imports System.Data.Common

Public Class Credit
    Inherits BLLBasic
    Implements IDisposable
    Private Language As New CableSoft.BLL.Language.SO61.AccountLanguage
    Private CitemName As String = Nothing
    Private UCCode As String = Nothing
    Private UCName As String = Nothing
    Private InstAddrNO As String = Nothing
    Private MduId As String = Nothing
    Private ServCode As String = Nothing
    Private ClctAreaCode As String = Nothing
    Private ClassCode1 As String = Nothing
    Private StrtCode As String = Nothing
    Private AreaCode As String = Nothing
    Private ClctEn As String = Nothing
    Private ClctName As String = Nothing
    
    Private _DAL As New CreditDAL(Me.LoginInfo.Provider)
    Private tbSO193 As DataTable = Nothing
    Private tbCD129 As DataTable = Nothing
    Private tbChangeResult As DataTable = Nothing
    Private lstbbTranspriority As List(Of Byte)
    Private totalBonus As Integer = 0
    Private totalSavepoint As Integer = 0
    Private Enum bbTranspriorityType
        HaveBonusStopDate = 1
        NoneBonusStopDate = 2
        SavePoint = 3
    End Enum
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
    Private Function QueryDefParaValue(ByVal transStoreCode As Integer) As Boolean
        Try
            tbChangeResult = New DataTable
            tbChangeResult.Columns.Add(New DataColumn("RowId", GetType(String)))
            tbChangeResult.Columns.Add(New DataColumn("MinusType", GetType(Integer)))
            tbChangeResult.Columns.Add(New DataColumn("CitemCode", GetType(Integer)))
            tbChangeResult.Columns.Add(New DataColumn("MinusPoint", GetType(Integer)))
            tbChangeResult.Columns.Add(New DataColumn("SeqNo", GetType(String)))

            Using dr As DbDataReader = DAO.ExecDtRdr(_DAL.QuerySO041)
                lstbbTranspriority = New List(Of Byte)
                While dr.Read
                    Dim strbbTranspriority As String = dr.Item(0).ToString
                    strbbTranspriority = strbbTranspriority.Replace(",", "").Replace(" ", "")
                    For i As Byte = 0 To strbbTranspriority.Length - 1
                        If i > 2 Then
                            Exit For
                        Else
                            lstbbTranspriority.Add(strbbTranspriority.Substring(i, 1))
                        End If
                    Next
                End While
            End Using
            Using dr As DbDataReader = DAO.ExecDtRdr(_DAL.QueryUCCode)
                While dr.Read
                    UCCode = dr.Item("CodeNo").ToString
                    UCName = dr.Item("Description").ToString
                End While
                dr.Dispose()
            End Using
            tbCD129 = DAO.ExecQry(_DAL.QueryCD129, New Object() {transStoreCode})
            If tbCD129.Rows.Count = 0 Then
                Throw New Exception("找不到轉換商城代碼檔！")
            End If
        Catch ex As Exception
            Throw
        End Try
        Return True
    End Function
    Private Function QueryBillDefaultValue(ByVal CitemCode As String, ByVal CustId As String,
                                           ByVal transStoreCode As Integer) As Boolean
        Try
            If Not String.IsNullOrEmpty(CitemCode) Then
                CitemName = DAO.ExecSclr(_DAL.QueryCitemName, New Object() {CitemCode}).ToString
            End If
           

           
            Using dr As DbDataReader = DAO.ExecDtRdr(_DAL.QueryCustId, New Object() {CustId, Me.LoginInfo.CompCode})
                While dr.Read
                    InstAddrNO = dr.Item("InstAddrNO")
                    MduId = dr.Item("MduId")
                    ServCode = dr.Item("ServCode")
                    ClctAreaCode = dr.Item("ClctAreaCode")
                    ClassCode1 = dr.Item("ClassCode1")
                End While
                dr.Dispose()
            End Using
            Using dr As DbDataReader = DAO.ExecDtRdr(_DAL.QueryAddrNo, New Object() {InstAddrNO, Me.LoginInfo.CompCode})
                While dr.Read
                    StrtCode = dr.Item("StrtCode")
                    AreaCode = dr.Item("AreaCode")
                    ClctEn = dr.Item("ClctEn")
                    ClctName = dr.Item("ClctName")
                End While
                dr.Dispose()
            End Using

        Catch ex As Exception
            Throw ex

        End Try
        Return True
    End Function
    Private Sub AddResultRecord(ByVal RowId As String, ByVal MinusType As Integer, _
                                            ByVal CitemCode As String, ByVal MinusValue As String,
                                            ByVal SeqNo As String)
        Try
            Dim rwNew As DataRow = tbChangeResult.NewRow
            With rwNew
                .Item("RowId") = RowId
                .Item("MinusType") = MinusType
                If Not String.IsNullOrEmpty(CitemCode) Then
                    .Item("CitemCode") = Integer.Parse(CitemCode)
                End If
                .Item("MinusPoint") = MinusValue
                .Item("SeqNo") = SeqNo
            End With
            tbChangeResult.Rows.Add(rwNew)
            tbChangeResult.AcceptChanges()
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Private Function CalculatePoint(ByVal TotalMinusPoint As Integer, ByVal MinusType As bbTranspriorityType) As Integer
        Try
            Select Case MinusType
                Case bbTranspriorityType.HaveBonusStopDate
                    For Each rw As DataRow In tbSO193.Rows

                        If (Not DBNull.Value.Equals(rw.Item("bonusStopDate"))) AndAlso
                            (DBNull.Value.Equals(rw.Item("closeBillno"))) Then
                            If (Date.Parse(rw.Item("bonusStopDate").ToString) > Date.Now) AndAlso
                                (Integer.Parse(rw.Item("bonus")) - Integer.Parse(rw.Item("Usedbonus")) > 0) Then
                                If (Integer.Parse(rw.Item("bonus")) - Integer.Parse(rw.Item("Usedbonus")) >= TotalMinusPoint) Then
                                    AddResultRecord(rw.Item("RowId"), 2, rw.Item("CitemCode").ToString, _
                                                    TotalMinusPoint, rw.Item("SeqNo"))
                                    Return 0
                                Else
                                    AddResultRecord(rw.Item("RowId"), 2, rw.Item("CitemCode").ToString,
                                                    Integer.Parse(rw.Item("bonus")) - Integer.Parse(rw.Item("Usedbonus")),
                                                    rw.Item("SeqNo"))
                                    TotalMinusPoint = TotalMinusPoint - (Integer.Parse(rw.Item("bonus")) - Integer.Parse(rw.Item("Usedbonus")))

                                End If
                            End If
                        End If
                    Next
                Case bbTranspriorityType.NoneBonusStopDate
                    For Each rw As DataRow In tbSO193.Rows
                        If (DBNull.Value.Equals(rw.Item("bonusStopDate"))) AndAlso (DBNull.Value.Equals(rw.Item("closeBillno"))) Then
                            If (Integer.Parse(rw.Item("bonus")) - Integer.Parse(rw.Item("Usedbonus")) > 0) Then
                                If (Integer.Parse(rw.Item("bonus")) - Integer.Parse(rw.Item("Usedbonus")) >= TotalMinusPoint) Then
                                    AddResultRecord(rw.Item("RowId"), 2, rw.Item("CitemCode").ToString, _
                                                    TotalMinusPoint, rw.Item("SeqNo"))
                                    Return 0
                                Else
                                    AddResultRecord(rw.Item("RowId"), 2, rw.Item("CitemCode").ToString,
                                                    Integer.Parse(rw.Item("bonus")) - Integer.Parse(rw.Item("Usedbonus")),
                                                    rw.Item("SeqNo"))
                                    TotalMinusPoint = TotalMinusPoint - (Integer.Parse(rw.Item("bonus")) - Integer.Parse(rw.Item("Usedbonus")))

                                End If
                            End If
                        End If
                    Next
                Case bbTranspriorityType.SavePoint
                    For Each rw As DataRow In tbSO193.Rows
                        If (Integer.Parse(rw.Item("Savepoint")) > 0) AndAlso (DBNull.Value.Equals(rw.Item("closeBillno"))) Then
                            If (Integer.Parse(rw.Item("Savepoint")) - Integer.Parse(rw.Item("UsedSavepoint")) > 0) Then
                                If (Integer.Parse(rw.Item("Savepoint")) - Integer.Parse(rw.Item("UsedSavepoint")) >= TotalMinusPoint) Then
                                    AddResultRecord(rw.Item("RowId"), 1, rw.Item("CitemCode").ToString, _
                                                    TotalMinusPoint, rw.Item("SeqNo"))
                                    Return 0
                                Else
                                    AddResultRecord(rw.Item("RowId"), 2, rw.Item("CitemCode").ToString,
                                                    Integer.Parse(rw.Item("Savepoint")) - Integer.Parse(rw.Item("UsedSavepoint")),
                                                    rw.Item("SeqNo"))
                                    TotalMinusPoint = TotalMinusPoint - (Integer.Parse(rw.Item("Savepoint")) - Integer.Parse(rw.Item("UsedSavepoint")))

                                End If
                            End If
                        End If
                    Next
                Case Else
                    Return TotalMinusPoint
            End Select
            Return TotalMinusPoint
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Private Function IsCanChange(ByVal MinusPoint As Integer) As Boolean
        Try
            If tbSO193.Rows.Count = 0 Then
                Return False
            End If
            '判斷扣點順序 1.有到期日紅利2.無到期日紅利3.儲值點數
            For Each ord As Integer In lstbbTranspriority
                If MinusPoint <= 0 Then
                    Exit For
                End If
                Select Case ord
                    Case 1
                        MinusPoint = CalculatePoint(MinusPoint, bbTranspriorityType.HaveBonusStopDate)
                    Case 2
                        MinusPoint = CalculatePoint(MinusPoint, bbTranspriorityType.NoneBonusStopDate)
                    Case 3
                        MinusPoint = CalculatePoint(MinusPoint, bbTranspriorityType.SavePoint)
                End Select
            Next
            Return MinusPoint = 0
        Catch ex As Exception
            Throw ex
        End Try
        Return True
    End Function
    Private Function InsSO033bbTrans(ByVal Transseqno As String, ByVal transStoreCode As String) As Boolean
        Dim transSavepoint As Integer = 0
        Dim transbonus As Integer = 0
        transbonus = 0
        transSavepoint = 0

        Try
            For Each rw As DataRow In tbChangeResult.Rows
                Select Case Integer.Parse(rw.Item("MinusType"))
                    Case 2
                        transbonus = transbonus + Integer.Parse(rw.Item("MinusPoint"))
                    Case 1
                        transSavepoint = transSavepoint + Integer.Parse(rw.Item("MinusPoint"))
                End Select
            Next
            'Return String.Format("Insert Into SO033BBTRANS (bbAccountID,SeqNo,transSavepoint," & _
            '                                "transbonus,transdate,transStoreCode,Rate,Mallpoint,Mallbonus,UpdTime,UpdEn) " & _
            '                                " Values ({0}0,{0}1,{0}2," & _
            '                                " {0}3,sysdate,{0}4,{0}5,{0}6,{0}7,sysdate,{0}8)", Sign)
            DAO.ExecNqry(_DAL.InsSO033BBTrans, New Object() {tbSO193.Rows(0).Item("bbAccountId"),
                                                             Transseqno, transSavepoint, transbonus, transStoreCode,
                                                             tbCD129.Rows.Item("Rate")})

        Catch ex As Exception
            Throw
        End Try
        Return True
    End Function
    Public Function ChangePoint(ByVal transStoreCode As Integer, ByVal bbAccountID As String,
                                ByVal FaciSeqNo As String, ByVal FaciSNo As String, ByVal MinusPoint As Integer) As RIAResult
        Dim result As New RIAResult()
        result.ResultBoolean = False
        Dim updNowDate As Date = Date.Now
        Dim cn As DbConnection = DAO.GetConn()
        Dim trans As DbTransaction = Nothing
        Dim blnAutoClose As Boolean = False
        Dim Transseqno As String = Nothing
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
        Try
            If Not QueryDefParaValue(transStoreCode) Then
                result.ResultBoolean = False
                result.ErrorCode = -1
                result.ErrorMessage = "QueryDefParaValue"
            End If
            tbSO193 = DAO.ExecQry(_DAL.QuerySO193, New Object() {bbAccountID})
            '過慮可以扣除的SO193資料
            If tbSO193.Rows.Count > 0 Then
                totalBonus = tbSO193.Rows(0).Item("TotalBonus")
                totalSavepoint = tbSO193.Rows(0).Item("TotalSavePoint")
                Select Case Integer.Parse(tbCD129.Rows(0).Item("Condition"))
                    '任何資料都可扣
                    Case 0
                        '只能扣除儲值點數
                    Case 1
                        Dim lstrw As List(Of DataRow) = tbSO193.AsEnumerable.Where(Function(rw As DataRow)
                                                                                       Return Integer.Parse(rw.Item("Savepoint")) <= 0
                                                                                   End Function).ToList()
                        For Each rw As DataRow In lstrw
                            tbSO193.Rows.Remove(rw)
                            tbSO193.AcceptChanges()
                        Next
                        '只能扣除紅利點數
                    Case 2
                        Dim lstrw As List(Of DataRow) = tbSO193.AsEnumerable.Where(Function(rw As DataRow)
                                                                                       Return Integer.Parse(rw.Item("bonus")) <= 0
                                                                                   End Function).ToList()
                        For Each rw As DataRow In lstrw
                            tbSO193.Rows.Remove(rw)
                            tbSO193.AcceptChanges()
                        Next
                End Select
                If Not IsCanChange(MinusPoint) Then
                    result.ErrorCode = -1
                    result.ErrorMessage = "點數不足！"
                    result.ResultBoolean = False
                    Return result
                End If
                Transseqno = DAO.ExecSclr(_DAL.GetTransseqno)
                If Not InsSO033bbTrans(Transseqno, transStoreCode) Then
                    result.ErrorCode = -2
                    result.ErrorMessage = "Insert SO033bbTrans 失敗"
                    result.ResultBoolean = False
                    Return result
                End If
            End If

            result.ResultBoolean = True
        Catch ex As Exception
            trans.Rollback()
            result.ResultBoolean = False
            result.ErrorCode = -99
            result.ErrorMessage = ex.ToString
        Finally
        End Try
        Return result
    End Function
#Region "IDisposable Support"
    Private disposedValue As Boolean ' 偵測多餘的呼叫

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: 處置 Managed 狀態 (Managed 物件)。
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
