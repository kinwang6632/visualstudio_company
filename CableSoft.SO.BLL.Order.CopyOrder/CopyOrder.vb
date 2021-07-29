Imports System.Data.Common
Imports CableSoft.BLL.Utility
Imports CableSoft.Utility.DataAccess

Public Class CopyOrder
    Inherits CableSoft.BLL.Utility.BLLBasic
    Implements IDisposable
    Private _DAL As New CopyOrderDALMultiDB(Me.LoginInfo.Provider)
    Private lang As CableSoft.BLL.Language.SO61.CopyOrderLanguage

    Public Sub New()

    End Sub
    Public Sub New(ByVal LoginInfo As LoginInfo)
        MyBase.New(LoginInfo)

    End Sub
    Public Sub New(ByVal LoginInfo As LoginInfo, ByVal DBConnection As System.Data.Common.DbConnection)
        MyBase.New(LoginInfo, DBConnection)

    End Sub
    Public Sub New(ByVal LoginInfo As LoginInfo, ByVal DAO As CableSoft.Utility.DataAccess.DAO)
        MyBase.New(LoginInfo, DAO)

    End Sub
    Public Function CanEdit(ByVal OrderNo As String) As RIAResult
        Dim obj As New CableSoft.SO.BLL.Order.Edit.OrderEdit(Me.LoginInfo, Me.DAO)

        Try
            Return obj.CanCopyOrder(OrderNo)
        Finally
            If obj IsNot Nothing Then
                obj.Dispose()
                obj = Nothing
            End If

        End Try
    End Function
    Public Function ChkReturnCode(ByVal OrderNo As String) As DataTable
       
        Return DAO.ExecQry(_DAL.ChkReturnCode, New Object() {OrderNo})
    End Function
    Public Function GetCustId(ByVal OrderNo As String) As DataTable
        Return DAO.ExecQry(_DAL.GetCustId, New Object() {OrderNo})
    End Function
    Public Function GetCloseWipData(ByVal OrderNo As String, ByVal CustId As Integer, _
                                    ByVal IncludeOrder As Boolean, ByVal WorkType As Integer) As RIAResult
        Dim r As New RIAResult(0, "")
        r.ResultXML = "0"
        r.ResultBoolean = True
        Try
            Using o As New CableSoft.SO.BLL.Wip.CloseWip.CallOK(Me.LoginInfo, Me.DAO)
                r.ResultDataSet = o.GetCloseWipData(OrderNo, CustId, IncludeOrder, WorkType).DataSet.Copy()
                o.Dispose()
            End Using
        Catch ex As Exception
            r.ResultBoolean = False
            r.ErrorMessage = ex.ToString
        End Try
        Return r
    End Function
    Public Function ChkAllSNOCitemCount(ByVal CustId As Integer, ByVal AllSNO As String) As RIAResult
        Dim r As New RIAResult(0, "")
        r.ResultXML = "0"
        r.ResultBoolean = True
        Try
            If String.IsNullOrEmpty(AllSNO) Then AllSNO = ""
            Using o As New CableSoft.SO.BLL.Wip.CloseWip.CallOK(Me.LoginInfo, Me.DAO)
                'If o.ChkAllSNOCitemCount(CustId, AllSNO) Then
                '    r.ResultXML = "1"
                'End If
                r = o.ChkAllSNOCitem(CustId, AllSNO)
                o.Dispose()
            End Using
        Catch ex As Exception
            r.ResultBoolean = False
            r.ErrorMessage = ex.ToString
        End Try
        Return r
    End Function
    Public Function Execute(ByVal OrderNo As String, ByVal AllSNO As String, ByVal WorkType As Integer,
                           ByVal ExecTab As DataTable, ByVal ShouldRegPriv As Boolean,
                          ByVal CustId As Integer, ByVal IsOrderTurnSend As Boolean, ByVal OtherTable As DataTable) As RIAResult
        Dim cn As DbConnection = DAO.GetConn()
        Dim trans As DbTransaction = Nothing
        Dim blnAutoClose As Boolean = False

        Dim r As New RIAResult(0, Nothing)
        r.ResultBoolean = True
        r.ErrorCode = 0
        r.ErrorMessage = Nothing
        If String.IsNullOrEmpty(AllSNO) Then AllSNO = ""
        If Me.DAO.Transaction IsNot Nothing Then
            trans = DAO.Transaction
        Else
            If cn IsNot Nothing AndAlso cn.State <> ConnectionState.Open Then
                cn.ConnectionString = Me.LoginInfo.ConnectionString
                cn.Open()
            End If
            trans = cn.BeginTransaction
            Me.DAO.Transaction = trans
            Me.DAO.AutoCloseConn = False
            blnAutoClose = True

        End If
        DAO.AutoCloseConn = False
        If blnAutoClose Then
            '  CableSoft.BLL.Utility.Utility.SetClientInfo(DAO, LoginInfo.EntryId, lang.ClientInfoString)
        End If

        'Dim o As New CableSoft.SO.BLL.Wip.CloseWip.Save(Me.LoginInfo, Me.DAO)

        If ExecTab IsNot Nothing Then
            'r = o.SaveData(AllSNO, WorkType, ExecTab, ShouldRegPriv, CustId, OrderNo, IsOrderTurnSend, OtherTable)

            Using o As New CableSoft.SO.BLL.Wip.CloseWip.Save(Me.LoginInfo, Me.DAO)
                r = o.SaveData(AllSNO, WorkType, ExecTab, ShouldRegPriv, CustId, OrderNo, IsOrderTurnSend, OtherTable)
            End Using
        End If
        If r.ResultBoolean Then
            r = Execute(OrderNo)
        End If

        If blnAutoClose Then
            If r.ResultBoolean Then
                trans.Commit()
            Else
                trans.Rollback()
            End If
        End If
        If blnAutoClose Then
            CableSoft.BLL.Utility.Utility.ClearClientInfo(DAO)
            If trans IsNot Nothing Then
                trans.Dispose()
                trans = Nothing
            End If
            If cn IsNot Nothing Then
                cn.Close()
                cn.Dispose()
                cn = Nothing
            End If
            DAO.AutoCloseConn = True
        End If

        Return r
    End Function
    Private Function SelectNoCopySO105Detail(ByVal OrderNo As String) As DataTable
        Dim dtRet As DataTable = Nothing
        Try
            dtRet = DAO.ExecQry(_DAL.NoCopySO105Detail, New Object() {OrderNo})
            dtRet.TableName = "NoCopy"
            Return dtRet.Copy
        Catch ex As Exception
            Throw ex
        Finally
            If dtRet IsNot Nothing Then
                dtRet.Dispose()
                dtRet = Nothing
            End If
        End Try
    End Function
    Private Function DeleteSO105Detail(ByVal NewOrder As DataTable, ByVal NoCopy As DataTable) As DataTable
        Dim lstRow As New List(Of DataRow)
        Dim dtRet As DataTable = NewOrder.Copy


        For Each row As DataRow In NoCopy.Rows
            Dim lst As List(Of DataRow) = NewOrder.AsEnumerable.Where(Function(ByVal newRow As DataRow)
                                                                          If newRow.Item("AutoSerialNo") = row.Item("AutoSerialNo") Then
                                                                              Return False
                                                                          End If
                                                                          Return True
                                                                      End Function).ToList
            If lst.Count > 0 Then
                lstRow.AddRange(lst.ToArray)
            End If

        Next
        If lstRow.Count > 0 Then
            dtRet.Rows.Clear()
            For Each rw As DataRow In lstRow
                dtRet.Rows.Add(rw.ItemArray)
            Next
        End If
        Return dtRet.Copy
    End Function
    Private Function DeleteSO105B(ByVal NewOrder As DataTable, ByVal NoCopy As DataTable) As DataTable
        Dim lstRow As New List(Of DataRow)
        Dim dtRet As DataTable = NewOrder.Copy
        'dtRet.Rows.Clear()

        For Each row As DataRow In NoCopy.Rows
            Dim lst As List(Of DataRow) = NewOrder.AsEnumerable.Where(Function(ByVal newRow As DataRow)
                                                                          If newRow.Item("WorkerType") = "I" AndAlso newRow.Item("CodeNo") = row.Item("CodeNo") Then
                                                                              Return False
                                                                          End If
                                                                          Return True
                                                                      End Function).ToList
            If lst.Count > 0 Then
                lstRow.AddRange(lst.ToArray)
            End If

        Next
        If lstRow.Count > 0 Then
            dtRet.Rows.Clear()
            For Each rw As DataRow In lstRow
                dtRet.Rows.Add(rw.ItemArray)
            Next
        End If
        Return dtRet.Copy
    End Function
    Public Function Execute(ByVal OrderNo As String) As RIAResult
        Dim cn As DbConnection = DAO.GetConn()
        Dim trans As DbTransaction = Nothing
        Dim blnAutoClose As Boolean = False
        Dim dtReturn As New DataTable("Order")
        Dim ds As New DataSet()
        Dim updDate As Date = DateTime.Now
        Dim aResvTime As Object = Nothing
        Dim dtNoCopy As DataTable = Nothing
        dtReturn.Columns.Add("CustId", GetType(Int32))
        dtReturn.Columns.Add("OrderNo", GetType(String))
        Dim NewOrderNo As String = ""
        Dim CustId As Integer = -1
        Dim Util As CableSoft.SO.BLL.Utility.Utility = Nothing
        Dim r As New RIAResult(0, Nothing)
        r.ResultBoolean = True
        r.ErrorCode = 0
        r.ErrorMessage = Nothing

        Try

            If DAO.Transaction IsNot Nothing Then
                trans = DAO.Transaction
                blnAutoClose = False
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


            Dim arraySQL As ArrayList = _DAL.GetOrder()
            Dim blnUpdNewTime As Boolean = False

            Util = New CableSoft.SO.BLL.Utility.Utility(LoginInfo, Me.DAO)
            NewOrderNo = Util.GetOrderNo()
            Dim dicFaciSeqNo As New Dictionary(Of String, String)
            CableSoft.BLL.Utility.Utility.SetClientInfo(Me.DAO, LoginInfo.EntryName)
            dtNoCopy = SelectNoCopySO105Detail(OrderNo)
            For Index As Integer = 0 To arraySQL.Count - 1
                blnUpdNewTime = False
                Dim aSql As String = arraySQL.Item(Index)
                Dim OldOrder As DataTable = DAO.ExecQry(aSql, New Object() {OrderNo}, False)
                Dim NewOrder As DataTable = OldOrder.Copy
                Dim TableName As String = ""
                Dim SeqName As String = ""
                Select Case Index
                    Case 0
                        TableName = "SO105"
                        blnUpdNewTime = True
                    Case 1
                        TableName = "SO105D1"
                    Case 2
                        TableName = "SO105B"
                        SeqName = "S_SO105B"
                        NewOrder = DeleteSO105B(NewOrder, dtNoCopy)
                    Case 3
                        TableName = "SO105C"
                        SeqName = "S_SO105C"
                        blnUpdNewTime = True
                    Case 4
                        TableName = "SO105D"
                        SeqName = "S_SO105D"
                    Case 5
                        TableName = "SO105DETAIL"
                        SeqName = "S_SO105DETAIL"
                        NewOrder = DeleteSO105Detail(NewOrder, dtNoCopy)
                    Case 6
                        TableName = "SO105I"
                        SeqName = "S_SO105I"
                        blnUpdNewTime = True
                    Case 7
                        TableName = "SO105J"
                        SeqName = String.Empty
                    Case 8
                        TableName = "SO105K"
                        SeqName = String.Empty
                End Select

                For RowIndex As Integer = 0 To NewOrder.Rows.Count - 1
                    Dim NewOrderRow As DataRow = NewOrder.Rows.Item(RowIndex)
                    With NewOrderRow
                        .Item("OrderNo") = NewOrderNo
                        '取得自動流水號欄位
                        If SeqName.Length > 0 Then
                            .Item("AutoSerialNo") = Integer.Parse(Util.GetSequenceNo(SeqName, 10))
                        End If
                        '單頭需回填來源訂單單號跟原始訂單單號
                        If Index = 0 Then
                            .Item("SourceOrderNo") = OldOrder.Rows(RowIndex).Item("OrderNo")
                            .Item("OldOrderNo") = OldOrder.Rows(RowIndex).Item("OldOrderNo")
                            If .IsNull("OldOrderNo") Then
                                .Item("OldOrderNo") = OldOrder.Rows(RowIndex).Item("OrderNo")
                            End If
                            '.Item("ResvTime") = DBNull.Value
                            .Item("ResvTime") = OldOrder.Rows(RowIndex).Item("ResvTime")
                            aResvTime = .Item("ResvTime")
                            .Item("ReturnDescCode") = DBNull.Value
                            .Item("ReturnDescName") = DBNull.Value
                            .Item("ReturnCode") = DBNull.Value
                            .Item("ReturnName") = DBNull.Value
                            .Item("FinTime") = DBNull.Value
                            .Item("UpdTime") = CableSoft.BLL.Utility.DateTimeUtility.GetDTString(updDate)
                            If blnUpdNewTime Then
                                .Item("NewUpdTime") = updDate
                            End If
                            .Item("UpdEn") = LoginInfo.EntryName
                            .Item("AcceptEn") = LoginInfo.EntryId
                            .Item("AcceptName") = LoginInfo.EntryName
                            .Item("AcceptTime") = updDate
                            '記錄CustId
                            CustId = Int32.Parse(.Item("CustId"))
                        ElseIf Index = 4 Then
                            If DBNull.Value.Equals(OldOrder.Rows(RowIndex).Item("FaciSeqNo")) OrElse
                                OldOrder.Rows(RowIndex).Item("FaciSeqNo").ToString.Length < 15 Then
                                .Item("FaciSeqNo") = Util.GetFaciSeqNo(True)
                                If Not DBNull.Value.Equals(OldOrder.Rows(RowIndex).Item("FaciSeqNo")) Then
                                    dicFaciSeqNo.Add(OldOrder.Rows(RowIndex).Item("FaciSeqNo"),
                                                     .Item("FaciSeqNo"))
                                End If
                            End If
                        ElseIf Index = 5 Then
                            '派工資料部份要將部份欄位清除
                            .Item("SNO") = DBNull.Value
                            '.Item("ResvTime") = DBNull.Value
                            .Item("ResvTime") = aResvTime
                            .Item("ReturnDescCode") = DBNull.Value
                            .Item("ReturnDescName") = DBNull.Value
                            .Item("ReturnCode") = DBNull.Value
                            .Item("ReturnName") = DBNull.Value
                            .Item("FinTime") = DBNull.Value
                            .Item("ReturnName") = DBNull.Value
                        End If
                    End With
                    Dim cmd As DbCommand = DAO._factory.CreateCommand()
                    cmd.Connection = cn
                    cmd.Transaction = trans
                    If Not Util.GetInsertCommand(NewOrder, TableName, RowIndex, cmd) Then
                        If blnAutoClose Then
                            trans.Rollback()
                        Else
                            r.ResultBoolean = False
                            r.ErrorCode = -1
                            r.ErrorMessage = TableName & " Insert Error "
                            Return r
                        End If
                    End If
                    cmd.ExecuteNonQuery()
                Next
            Next
            'SO105B,SO105C,SO105J,SO105K 需將有對應該 FaciSeqNo 從舊值改為新值 For Jacky By Kin 2013/07/18
            For i As Int32 = 0 To dicFaciSeqNo.Count - 1
                For index As Int32 = 0 To 3
                    Select Case index
                        Case 0
                            DAO.ExecNqry(_DAL.UpdateFaciSeqNo("SO105B"),
                                         New Object() {dicFaciSeqNo.Values(i), NewOrderNo, dicFaciSeqNo.Keys(i)})
                        Case 1
                            DAO.ExecNqry(_DAL.UpdateFaciSeqNo("SO105C"),
                                         New Object() {dicFaciSeqNo.Values(i), NewOrderNo, dicFaciSeqNo.Keys(i)})
                        Case 2
                            DAO.ExecNqry(_DAL.UpdateFaciSeqNo("SO105J"),
                                         New Object() {dicFaciSeqNo.Values(i), NewOrderNo, dicFaciSeqNo.Keys(i)})
                        Case 3
                            DAO.ExecNqry(_DAL.UpdateFaciSeqNo("SO105K"),
                                         New Object() {dicFaciSeqNo.Values(i), NewOrderNo, dicFaciSeqNo.Keys(i)})
                    End Select
                Next


            Next


            Dim rwNew As DataRow = dtReturn.NewRow
            rwNew.Item("Custid") = CustId
            rwNew.Item("OrderNo") = NewOrderNo
            dtReturn.Rows.Add(rwNew)
            ds.Tables.Add(dtReturn)
            r.ResultDataSet = ds.Copy
            If blnAutoClose Then
                trans.Commit()
            End If
            'Util.Dispose()
        Catch ex As Exception
            If blnAutoClose Then
                'trans.Rollback()
            End If
            r.ErrorMessage = ex.ToString
            r.ErrorCode = -1
            r.ResultBoolean = False

        Finally
            If blnAutoClose Then
                If trans IsNot Nothing Then
                    trans.Dispose()
                    trans = Nothing
                End If
                If cn IsNot Nothing Then
                    cn.Close()
                    cn.Dispose()
                    cn = Nothing
                End If
                DAO.AutoCloseConn = True
            End If
            If Util IsNot Nothing Then
                Util.Dispose()
            End If
            If dtNoCopy IsNot Nothing Then
                dtNoCopy.Dispose()
                dtNoCopy = Nothing
            End If
            If dtReturn IsNot Nothing Then
                dtReturn.Dispose()
                dtReturn = Nothing
            End If
            If ds IsNot Nothing Then
                ds.Dispose()
                ds = Nothing
            End If
        End Try

        'aRet.ResultXML = NewOrderNo

        'Return aRet

        Return r
    End Function

#Region "IDisposable Support"
    Private disposedValue As Boolean ' 偵測多餘的呼叫

    ' IDisposable
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: 處置 Managed 狀態 (Managed 物件)。
                If (Me.MustDispose) AndAlso (Me.DAO IsNot Nothing) Then
                    DAO.Dispose()
                End If
                If _DAL IsNot Nothing Then
                    _DAL.Dispose()
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
