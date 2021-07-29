Imports CableSoft.BLL.Utility

Public Class BillingAPI712
    Inherits BLLBasic
    Implements IDisposable, CableSoft.BLL.BillingAPI.IBillingAPI
    Private Language As New CableSoft.BLL.Language.SO61.BillingAPI712Language
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
    Public Function Execute(SeqNo As Integer, InData As System.Data.DataSet) As CableSoft.BLL.Utility.RIAResult Implements CableSoft.BLL.BillingAPI.IBillingAPI.Execute
        Dim nTable As New DataTable() With {.TableName = "Main"}
        For Each table As DataTable In InData.Tables
            If table.TableName.ToLower = "settingdata".ToLower Then
                Continue For
            End If
            For Each col As DataColumn In table.Columns
                If nTable.Columns.Contains(col.ColumnName) = False Then
                    nTable.Columns.Add(New DataColumn() With {.ColumnName = col.ColumnName, .DataType = col.DataType})
                End If
            Next
        Next
        Using bll As New CableSoft.BLL.BillingAPI.BillingAPI(LoginInfo, DAO)
            Dim Count As Integer = InData.Tables("SNO").Rows.Count - 1
            Dim SNOs(Count) As String
            Dim Amounts(Count) As String
            Dim nRow As DataRow = nTable.NewRow()
            nTable.Rows.Add(nRow)
            For intLoop As Integer = 0 To Count
                Dim Row As DataRow = InData.Tables("SNO").Rows(intLoop)
                For Each col As DataColumn In InData.Tables("Main").Columns
                    nRow.Item(col.ColumnName) = InData.Tables("Main").Rows(0).Item(col.ColumnName)
                Next
                For Each col As DataColumn In InData.Tables("SNO").Columns
                    nRow.Item(col.ColumnName) = Row.Item(col.ColumnName)
                Next
                nRow.Item("APIID") = 235
                Dim nDataSet As New DataSet()
                nDataSet.Merge(nTable)
                Dim jsonStr As String = JsonServer.ToJson(nDataSet, JsonServer.JsonFormatting.None, JsonServer.NullValueHandling.Include, False, False)
                Dim result As RIAResult = bll.Execute(jsonStr, False)
                If result.ResultBoolean = False Then
                    result.ErrorMessage = String.Format(Language.retMessage, nRow.Item("SNo"), result.ErrorMessage)
                    Return result
                End If
                Amounts(intLoop) = result.ResultDataSet.Tables("AMT").Rows(0).Item(0)
                SNOs(intLoop) = nRow.Item("SNo")
            Next
            Return New RIAResult With {.ResultBoolean = True, .ResultDataSet = GetReturnData(SNOs, Amounts)}
        End Using
    End Function
    Private Function GetReturnData(SNOs() As String, Amounts() As String) As DataSet
        Dim RetData As New DataSet With {.DataSetName = "DataSet"}
        Dim RetTable As New DataTable With {.TableName = "SNO"}
        RetTable.Columns.Add(New DataColumn With {.ColumnName = "SNO", .DataType = GetType(String)})
        RetTable.Columns.Add(New DataColumn With {.ColumnName = "Amount", .DataType = GetType(String)})
        For intLoop As Integer = 0 To SNOs.GetUpperBound(0)
            Dim nRow As DataRow = RetTable.NewRow()
            RetTable.Rows.Add(nRow)
            nRow.Item("SNO") = SNOs(intLoop)
            nRow.Item("Amount") = Amounts(intLoop)
        Next
        RetData.Tables.Add(RetTable)
        Return RetData
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
