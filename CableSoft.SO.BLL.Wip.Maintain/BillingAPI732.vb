Imports CableSoft.BLL.Utility

Public Class BillingAPI732
    Inherits BLLBasic
    Implements IDisposable, CableSoft.BLL.BillingAPI.IBillingAPI

    Const aryChgColName As String = "WorkerEn,SignEn,SignDate,SatiCode,SatiName,BillStatus,ClctFlag"
    Private Language As New CableSoft.BLL.Language.SO61.BillingAPI732Language
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
            Dim snoTable As New DataTable() With {.TableName = "SNO"}
            Dim lstChgColumns As IList(Of DataColumn) = New List(Of DataColumn)
            For Each strColname As String In aryChgColName.Split(",")
                If InData.Tables("Main").Columns.Contains(strColname) Then
                    Using newCol As DataColumn = New DataColumn()
                        newCol.DataType = InData.Tables("Main").Columns(strColname).DataType
                        newCol.ColumnName = InData.Tables("Main").Columns(strColname).ColumnName
                        newCol.DefaultValue = InData.Tables("Main").Rows(0).Item(strColname)
                        lstChgColumns.Add(newCol)
                    End Using
                End If
            Next
            For Each inTable As DataTable In InData.Tables
                For i As Integer = 0 To lstChgColumns.Count - 1
                    Select Case inTable.TableName.ToUpper
                        Case "main".ToUpper
                            inTable.Columns.Remove(lstChgColumns(i).ColumnName)
                        Case "sno".ToUpper
                            inTable.Columns.Add(lstChgColumns(i))
                    End Select
                Next
            Next
            Using bll As New CableSoft.BLL.BillingAPI.BillingAPI(LoginInfo, DAO)

            Dim Count As Integer = InData.Tables("SNO").Rows.Count - 1

            InData.Tables("Main").Rows(0).Item("APIID") = 222
            For intLoop As Integer = 0 To Count
                Using snoTableClone As DataTable = InData.Tables("Sno").Clone


                    snoTableClone.Rows.Add(InData.Tables("Sno").Rows(intLoop).ItemArray)
                    'For Each col As DataColumn In snoTable.Columns
                    '    nRow.Item(col.ColumnName) = Row.Item(col.ColumnName)
                    'Next

                    Using nDataSet As New DataSet()
                        nDataSet.Merge(InData.Tables("main").Copy)
                        nDataSet.Merge(snoTableClone)
                        Dim jsonStr As String = JsonServer.ToJson(nDataSet, JsonServer.JsonFormatting.None, JsonServer.NullValueHandling.Include, False, False)
                        Dim result As RIAResult = bll.Execute(jsonStr, False)
                        If result.ResultBoolean = False Then
                            result.ErrorMessage = String.Format(Language.retMessage, snoTableClone.Rows(0).Item("SNo"), result.ErrorMessage)
                            Return result
                        End If
                    End Using

                End Using

            Next
            If lstChgColumns.Count > 0 Then
                For Each col As DataColumn In lstChgColumns
                    col.Dispose()
                    col = Nothing
                Next
                lstChgColumns.Clear()
            End If
            Return New RIAResult With {.ResultBoolean = True, .ErrorMessage = Nothing}
            End Using

     
      




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
